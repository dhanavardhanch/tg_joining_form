/**
 * TrooGood Employee Onboarding → Google Drive PDF & Multi-Tab Sheet Automation
 * ============================================================================
 * Official Backend Script for TrooGood / Mformillet Foods Pvt Ltd
 * Target Sheet: https://docs.google.com/spreadsheets/d/1F9_-9uE4CHwJUFEf73XuaApEMVYhGLPMc1KNe4g4uVM/edit
 * Target Drive: https://drive.google.com/drive/folders/1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t
 *
 * Capabilities:
 *  1. Auto-increments Reference ID in sequential format: TG-00001, TG-00002...
 *     (Resets automatically if test rows are deleted).
 *  2. Creates employee subfolder: "TG-00001 - Candidate Name" (Option B).
 *  3. Saves Full Original HD Photograph, Aadhaar Card document, and Digital Signature.
 *  4. Generates branded Joining Form PDF with top-left photo.
 *  5. If candidate requests PF Exemption, automatically generates a dedicated
 *     statutory "PF Exemption Undertaking & Self-Declaration" PDF signed by employee.
 *  6. Logs quick overview into Tab 1: "Submissions" (Exact 25 columns).
 *  7. Logs every single form field into Tab 2: "Complete form data" (All fields).
 *  8. Provides real-time duplicate checking endpoint (by Mobile / Aadhaar).
 *  9. Emails HR notification to sunaina@troogood.com.
 */

// Target Configuration
var SPREADSHEET_ID    = '1F9_-9uE4CHwJUFEf73XuaApEMVYhGLPMc1KNe4g4uVM';
var PARENT_FOLDER_ID  = '1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t';
var HR_EMAIL          = 'sunaina@troogood.com';
var TAB_SUBMISSIONS   = 'Submissions';
var TAB_COMPLETE_DATA = 'Complete form data';

/** Helper to get the target spreadsheet */
function getSpreadsheet() {
  if (SPREADSHEET_ID) {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch(e) {}
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Web App GET endpoint:
 *  - Default: Health check
 *  - action=checkDuplicate&mobile=XXXXX&aadhaar=YYYYY: Instant duplicate lookup (< 50ms)
 *  - action=getNextId: Preview next available sequential ID
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  
  if (action === 'checkDuplicate') {
    var mobile = (e.parameter.mobile || '').trim();
    var aadhaar = (e.parameter.aadhaar || '').trim();
    var dup = checkDuplicateCandidate(mobile, aadhaar);
    return json(dup);
  }

  if (action === 'getNextId') {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(TAB_SUBMISSIONS);
    var nextId = getNextReferenceId(sheet);
    return json({ ok: true, nextId: nextId });
  }

  return json({
    ok: true,
    message: 'TrooGood Onboarding Drive & Sheet Endpoint is live. Brand color: #00B5E8.'
  });
}

/**
 * Web App POST endpoint:
 * Processes onboarding submission from TrooGood Joining Form.
 */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = getSpreadsheet();
    if (!ss) throw new Error('Spreadsheet could not be accessed. Verify SPREADSHEET_ID.');

    var subSheet = ss.getSheetByName(TAB_SUBMISSIONS) || ss.insertSheet(TAB_SUBMISSIONS);
    var compSheet = ss.getSheetByName(TAB_COMPLETE_DATA) || ss.insertSheet(TAB_COMPLETE_DATA);

    // 1. Duplicate Verification Check
    var cleanMobile = String(data.mobile || '').replace(/\D/g, '').slice(-10);
    var cleanAadhaar = String(data.aadhaar || '').replace(/\D/g, '').slice(-12);
    var dupCheck = checkDuplicateCandidate(cleanMobile, cleanAadhaar);
    if (dupCheck.exists) {
      return json({
        ok: false,
        error: 'DUPLICATE_CANDIDATE',
        message: 'This candidate is already registered under Reference ID: ' + dupCheck.reference + ' (' + dupCheck.name + '). Duplicate submissions are blocked.',
        reference: dupCheck.reference
      });
    }

    // 2. Sequential Reference ID (TG-00001, TG-00002, ...)
    var newRef = getNextReferenceId(subSheet);
    data.reference = newRef;
    var empName = (data.name || 'New Employee').trim();

    // 3. Locate Parent Google Drive Folder
    var parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    } catch(fErr) {
      var folders = DriveApp.getFoldersByName("TrooGood Employee Onboarding Documents");
      parentFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder("TrooGood Employee Onboarding Documents");
    }

    // 4. Create Employee Subfolder: "TG-00001 - Candidate Name" (Option B)
    var folderTitle = newRef + ' - ' + empName;
    var subFolder = parentFolder.createFolder(folderTitle);
    try {
      subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}

    // 5. Save Full HD Uploads into Subfolder
    var photoUrl = saveBase64File(subFolder, data.photoImage, 'Photograph - ' + empName + '.jpg', 'image/jpeg');
    data.photoLink = photoUrl;

    var aadhaarUrl = saveBase64File(subFolder, data.aadhaarImage, 'Aadhaar - ' + empName + '.jpg', 'image/jpeg');
    data.aadhaarLink = aadhaarUrl;

    var sigUrl = saveBase64File(subFolder, data.signatureImage, 'Signature - ' + empName + '.png', 'image/png');
    data.signatureLink = sigUrl;

    // 6. Generate Branded Joining Form PDF
    var pdfFile = createEmployeePdf(data, subFolder, empName);
    var pdfUrl = pdfFile ? pdfFile.getUrl() : '';
    data.pdfLink = pdfUrl;

    // 7. Dedicated PF Exemption Declaration PDF (if candidate selected the exemption)
    var pfPdfUrl = '';
    if (data.declarations && data.declarations.dec_pf_exclude) {
      var pfFile = createPfExemptionPdf(data, subFolder, empName);
      pfPdfUrl = pfFile ? pfFile.getUrl() : '';
    }
    data.pfExemptionPdfLink = pfPdfUrl;

    // 8. Log into Tab 1: "Submissions" (Exact 25 columns)
    logToSubmissionsSheet(subSheet, data, subFolder.getUrl(), pdfUrl, photoUrl, aadhaarUrl);

    // 9. Log into Tab 2: "Complete form data" (Comprehensive all fields)
    logToCompleteDataSheet(compSheet, data, subFolder.getUrl(), pdfUrl, pfPdfUrl, photoUrl, aadhaarUrl, sigUrl);

    // 10. Optional HR Email Notification
    notify(data, subFolder.getUrl(), pdfUrl, pfPdfUrl);

    return json({
      ok: true,
      reference: newRef,
      folderUrl: subFolder.getUrl(),
      pdfUrl: pdfUrl,
      pfExemptionPdfUrl: pfPdfUrl,
      photoUrl: photoUrl,
      aadhaarUrl: aadhaarUrl
    });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/**
 * Calculates next Reference ID in sequential format: TG-00001, TG-00002...
 * Scans Column B of "Submissions". If rows were deleted during testing,
 * it resets automatically based on the last remaining value (or restarts at TG-00001).
 */
function getNextReferenceId(sheet) {
  if (!sheet) return 'TG-00001';
  var lastRow = sheet.getLastRow();
  var maxNum = 0;

  if (lastRow > 1) {
    var refValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var i = 0; i < refValues.length; i++) {
      var cellVal = String(refValues[i][0] || '').trim();
      var match = cellVal.match(/^TG-(\d+)$/i);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxNum) maxNum = num;
      }
    }
  }

  var nextNum = maxNum + 1;
  return 'TG-' + ('00000' + nextNum).slice(-5);
}

/**
 * High-speed in-memory duplicate candidate lookup.
 * Checks Mobile Number (last 10 digits) and Aadhaar Number (last 12 digits).
 */
function checkDuplicateCandidate(mobile, aadhaar) {
  var ss = getSpreadsheet();
  if (!ss) return { exists: false };
  var sheet = ss.getSheetByName(TAB_SUBMISSIONS);
  if (!sheet) return { exists: false };

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { exists: false };

  var cleanMobile = (mobile || '').toString().replace(/\D/g, '').slice(-10);
  var cleanAadhaar = (aadhaar || '').toString().replace(/\D/g, '').slice(-12);

  if (cleanMobile.length < 10 && cleanAadhaar.length < 12) {
    return { exists: false };
  }

  // Batch read first 9 columns: Timestamp(A), Ref(B), Name(C), Unit(D), Desig(E), Salary(F), Mobile(G), DOJ(H), Aadhaar(I)
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  for (var i = 0; i < dataRange.length; i++) {
    var row = dataRange[i];
    var rowRef = String(row[1] || '').trim();
    var rowName = String(row[2] || '').trim();
    var rowMobile = String(row[6] || '').replace(/\D/g, '').slice(-10);
    var rowAadhaar = String(row[8] || '').replace(/\D/g, '').slice(-12);
    var rowDate = row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Kolkata', 'dd MMM yyyy') : '';

    var matchMob = (cleanMobile.length === 10 && rowMobile === cleanMobile);
    var matchAad = (cleanAadhaar.length === 12 && rowAadhaar === cleanAadhaar);

    if (matchMob || matchAad) {
      return {
        exists: true,
        reference: rowRef,
        name: rowName,
        date: rowDate,
        matchedField: matchMob ? 'mobile' : 'aadhaar'
      };
    }
  }

  return { exists: false };
}

/** Saves base64 string directly as high-resolution original file */
function saveBase64File(folder, base64Data, filename, mimeType) {
  if (!base64Data || base64Data.indexOf('base64,') === -1) return '';
  try {
    var rawBytes = Utilities.base64Decode(base64Data.split('base64,')[1]);
    var file = folder.createFile(Utilities.newBlob(rawBytes, mimeType, filename));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
    return file.getUrl();
  } catch(e) {
    return '';
  }
}

/** Formats text to preserve leading zeros in Sheets */
function textCell(val) {
  if (val === null || val === undefined) return '';
  var str = val.toString().trim();
  if (!str) return '';
  return str.indexOf("'") === 0 ? str : "'" + str;
}

/**
 * Tab 1: Logs into "Submissions" across the exact 25 columns
 */
function logToSubmissionsSheet(sheet, data, folderUrl, pdfUrl, photoUrl, aadhaarUrl) {
  var headers = [
    'Timestamp',               // A (1)
    'Reference ID',            // B (2)
    'Name',                    // C (3)
    'Unit',                    // D (4)
    'Designation',             // E (5)
    'Salary',                  // F (6)
    'Mobile Number',           // G (7)
    'Date of Joining',         // H (8)
    'Aadhaar Number',          // I (9)
    'PAN Number',              // J (10)
    'UAN Number',              // K (11)
    'ESI Number',              // L (12)
    'Health Insurance Number', // M (13)
    'Account Number',          // N (14)
    'IFSC Code',               // O (15)
    'Drive Folder Link',       // P (16)
    'PDF Application Link',    // Q (17)
    'Passport Photo Link',     // R (18)
    'Aadhaar Card Link',       // S (19)
    'Date of Birth',           // T (20)
    'Shift Timing',            // U (21)
    'Employee Type',           // V (22)
    'Contractor Name',         // W (23)
    'Employee Code',           // X (24)
    'Gender'                   // Y (25)
  ];

  // Set Row 1 headers with TrooGood Cyan style
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00B5E8')
    .setFontColor('#ffffff');

  var newRow = [
    new Date(),                                           // A: Timestamp
    data.reference || '',                                 // B: Reference ID
    data.name || '',                                      // C: Name
    data.unit || '',                                      // D: Unit
    data.desig || '',                                     // E: Designation
    data.salary || 'N/A',                                 // F: Salary
    textCell(data.mobile),                                // G: Mobile Number
    data.doj || '',                                       // H: Date of Joining
    textCell(data.aadhaar),                               // I: Aadhaar Number
    data.pan || '',                                       // J: PAN Number
    data.uan || 'N/A',                                    // K: UAN Number
    textCell(data.esic || 'N/A'),                         // L: ESI Number
    data.health_ins_no || 'N/A',                          // M: Health Insurance Number
    textCell(data.account),                               // N: Account Number
    data.ifsc || '',                                      // O: IFSC Code
    folderUrl || '',                                      // P: Drive Folder Link
    pdfUrl || '',                                         // Q: PDF Application Link
    photoUrl || '',                                       // R: Passport Photo Link
    aadhaarUrl || '',                                     // S: Aadhaar Card Link
    data.dob || '',                                       // T: Date of Birth
    data.shift || 'General Shift',                        // U: Shift Timing
    data.engagement || 'On TrooGood Rolls',               // V: Employee Type
    data.vendor || 'N/A',                                 // W: Contractor Name
    data.code || 'Auto-allotted',                         // X: Employee Code
    data.gender || 'Not specified'                        // Y: Gender
  ];

  sheet.appendRow(newRow);
}

/**
 * Tab 2: Logs into "Complete form data" storing every single captured field
 */
function logToCompleteDataSheet(sheet, data, folderUrl, pdfUrl, pfPdfUrl, photoUrl, aadhaarUrl, sigUrl) {
  var headers = [
    'Timestamp',
    'Reference ID',
    'Name',
    'Employee Code',
    'Unit',
    'Designation',
    'Salary',
    'Date of Joining',
    'Reporting Supervisor',
    'Shift Timing',
    'Employee Type',
    'Contractor Name',
    'Father / Husband Name',
    'Date of Birth',
    'Gender',
    'Marital Status',
    'Blood Group',
    'Languages Known',
    'Primary Mobile',
    'Alternate Contact',
    'Email Address',
    'Present Address',
    'Permanent Address',
    'Emergency Contact Name',
    'Emergency Relationship',
    'Emergency Mobile',
    'Aadhaar Number',
    'PAN Number',
    'Bank Name & Branch',
    'Account Number',
    'IFSC Code',
    'Documents Submitted',
    'Worked Before',
    'Previous Employer',
    'Previous Last Working Day',
    'Prior EPF Account',
    'UAN Number',
    'Previous PF Member ID',
    'Previous Establishment',
    'Enrolled in EPS 1995',
    'ESIC Insurance Number',
    'Health Insurance Covered',
    'Health Insurance Number',
    'International Worker',
    'Primary Nominee Name',
    'Primary Nominee Relationship',
    'Primary Nominee DOB',
    'Family Dependents Details',
    'Declarations Accepted Count',
    'PF Exemption Requested',
    'Signed By Name',
    'Signature Place',
    'Signature Timestamp',
    'Drive Folder Link',
    'Joining Form PDF Link',
    'PF Exemption PDF Link',
    'Passport Photo Link',
    'Aadhaar Card Link',
    'Signature Link'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#0096c2')
    .setFontColor('#ffffff');

  // Format all family dependents as single readable text
  var famStr = 'None';
  if (data.family && data.family.length > 0) {
    famStr = data.family.map(function(f, idx) {
      return (idx + 1) + '. ' + (f.name || '') + ' (' + (f.relation || '') + ')' + (f.dob ? ' DOB: ' + f.dob : '');
    }).join('; ');
  }

  // Count accepted declarations
  var declCount = 0;
  if (data.declarations) {
    for (var k in data.declarations) {
      if (data.declarations[k]) declCount++;
    }
  }

  var isPfExempt = (data.declarations && data.declarations.dec_pf_exclude) ? 'YES' : 'NO';

  var row = [
    new Date(),
    data.reference || '',
    data.name || '',
    data.code || 'Auto-allotted',
    data.unit || '',
    data.desig || '',
    data.salary || 'N/A',
    data.doj || '',
    data.reporting || 'N/A',
    data.shift || 'General Shift',
    data.engagement || 'On TrooGood Rolls',
    data.vendor || 'N/A',
    data.father || '',
    data.dob || '',
    data.gender || 'Not specified',
    data.marital || '',
    data.blood || 'N/A',
    data.languages || 'N/A',
    textCell(data.mobile),
    textCell(data.altmobile || 'N/A'),
    data.email || 'N/A',
    data.addr_present || '',
    data.addr_permanent || 'Same as present address',
    data.emg_name || '',
    data.emg_rel || '',
    textCell(data.emg_mobile),
    textCell(data.aadhaar),
    data.pan || 'N/A',
    data.bank || '',
    textCell(data.account),
    data.ifsc || '',
    data.docs || 'None',
    data.worked || 'No',
    data.prev_employer || 'N/A',
    data.prev_last_day || 'N/A',
    data.epf || 'No',
    data.uan || 'N/A',
    data.pf_account || 'N/A',
    data.prev_estb || 'N/A',
    data.eps || 'N/A',
    textCell(data.esic || 'N/A'),
    data.health_ins || 'No',
    data.health_ins_no || 'N/A',
    data.intl || 'No',
    data.nom_name || '',
    data.nom_rel || '',
    data.nom_dob || '',
    famStr,
    declCount + ' of 5',
    isPfExempt,
    data.sig_name || data.name || '',
    data.sig_place || 'Hyderabad',
    data.signedAt || new Date().toISOString(),
    folderUrl || '',
    pdfUrl || '',
    pfPdfUrl || 'N/A',
    photoUrl || '',
    aadhaarUrl || '',
    sigUrl || ''
  ];

  sheet.appendRow(row);
}

/**
 * Builds and saves the complete formatted Joining Form PDF
 */
function createEmployeePdf(data, subFolder, empName) {
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMMM yyyy, hh:mm a');
  
  var familyRows = '';
  if (data.family && data.family.length > 0) {
    for (var i = 0; i < data.family.length; i++) {
      var fam = data.family[i];
      familyRows += '<tr><td class="lbl">Family Member #' + (i + 1) + '</td><td>' +
        esc(fam.name || '') + (fam.relation ? ' (' + esc(fam.relation) + ')' : '') +
        (fam.dob ? ' · DOB: ' + esc(fam.dob) : '') + '</td></tr>';
    }
  } else {
    familyRows = '<tr><td class="lbl">Family Members</td><td>None declared</td></tr>';
  }

  var declCount = 0;
  if (data.declarations) {
    for (var k in data.declarations) {
      if (data.declarations[k]) declCount++;
    }
  }

  var photoHtml = '';
  if (data.photoImage && data.photoImage.indexOf('base64,') > -1) {
    photoHtml = '<img src="' + data.photoImage + '" style="width:96px; height:120px; object-fit:cover; border:2px solid #00B5E8; border-radius:4px; display:block;" alt="Photo" />';
  } else {
    photoHtml = '<div style="width:96px; height:120px; border:1.5px dashed #94a3b8; border-radius:4px; text-align:center; font-size:10px; color:#64748b; background:#f8fafc; padding-top:45px; box-sizing:border-box;">Affix Photo</div>';
  }

  var aadhaarDocHtml = '';
  if (data.aadhaarImage && data.aadhaarImage.indexOf('base64,') > -1) {
    aadhaarDocHtml = '<tr><td class="lbl">Aadhaar Card Document</td><td>' +
      '<img src="' + data.aadhaarImage + '" style="max-width:320px; max-height:190px; object-fit:contain; border:1.5px solid #00B5E8; border-radius:4px; display:block; margin-top:4px;" />' +
      '</td></tr>';
  }

  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    'body { font-family: Helvetica, Arial, sans-serif; color: #0f172a; margin: 0; padding: 24px; font-size: 12px; line-height: 1.5; background: #ffffff; }' +
    '.header-table { width: 100%; border-bottom: 2.5px solid #00B5E8; padding-bottom: 14px; margin-bottom: 18px; }' +
    '.brand { font-size: 13px; font-weight: 800; color: #00B5E8; text-transform: uppercase; letter-spacing: 0.08em; }' +
    'h1 { margin: 4px 0 2px 0; font-size: 20px; color: #0f172a; }' +
    '.meta { font-size: 11px; color: #64748b; margin-top: 4px; line-height: 1.5; }' +
    '.card { border: 1px solid #cbd5e1; border-radius: 6px; margin-bottom: 16px; overflow: hidden; }' +
    '.card-title { background: #e6f8fd; color: #0096c2; font-size: 12.5px; font-weight: bold; padding: 7px 12px; border-bottom: 1px solid #bcecf8; }' +
    '.table { width: 100%; border-collapse: collapse; }' +
    '.table td { padding: 6px 10px; border-bottom: 1px solid #f1f5f9; font-size: 11.5px; vertical-align: top; }' +
    '.table tr:last-child td { border-bottom: none; }' +
    '.table td.lbl { width: 34%; font-weight: bold; color: #475569; background: #f8fafc; border-right: 1px solid #f1f5f9; }' +
    '.sig-img { max-width: 260px; max-height: 85px; display: block; border-bottom: 1.5px solid #0f172a; margin-top: 6px; }' +
    '.footer { margin-top: 24px; padding-top: 10px; border-top: 1px dashed #cbd5e1; font-size: 10px; color: #94a3b8; text-align: center; }' +
    '</style></head><body>' +

    '<table class="header-table" style="border-collapse:collapse;">' +
      '<tr>' +
        '<td style="width:110px; vertical-align:top; padding-right:16px;">' +
          photoHtml +
        '</td>' +
        '<td style="vertical-align:top;">' +
          '<div class="brand">TROOGOOD · Mformillet Foods Pvt Ltd</div>' +
          '<h1>Employee Onboarding & Joining Form</h1>' +
          '<div class="meta">' +
            'Reference ID: <b>' + esc(data.reference || '') + '</b><br>' +
            'Candidate Name: <b>' + esc(empName) + '</b><br>' +
            'Unit: <b>' + esc(data.unit || '—') + '</b> | Role: <b>' + esc(data.desig || '—') + '</b><br>' +
            'Date of Submission: ' + todayStr +
          '</div>' +
        '</td>' +
      '</tr>' +
    '</table>' +

    // 1. Job Details
    '<div class="card">' +
      '<div class="card-title">1. Job & Position Details</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Employee Code</td><td>' + esc(data.code || 'Auto-allotted upon approval') + '</td></tr>' +
        '<tr><td class="lbl">Date of Joining</td><td>' + esc(data.doj || '') + '</td></tr>' +
        '<tr><td class="lbl">Unit / Location</td><td>' + esc(data.unit || '') + '</td></tr>' +
        '<tr><td class="lbl">Role / Designation</td><td>' + esc(data.desig || '') + '</td></tr>' +
        '<tr><td class="lbl">Monthly Gross Salary</td><td>' + esc(data.salary ? '₹' + data.salary : 'As per appointment order') + '</td></tr>' +
        '<tr><td class="lbl">Reporting Supervisor</td><td>' + esc(data.reporting || '') + '</td></tr>' +
        '<tr><td class="lbl">Shift Timing</td><td>' + esc(data.shift || 'General Shift') + '</td></tr>' +
        '<tr><td class="lbl">Employee Type</td><td>' + esc(data.engagement || 'On TrooGood Rolls') + (data.vendor ? ' (Contractor: ' + esc(data.vendor) + ')' : '') + '</td></tr>' +
      '</table>' +
    '</div>' +

    // 2. Personal Information
    '<div class="card">' +
      '<div class="card-title">2. Personal & Contact Information</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Full Name (as per Aadhaar)</td><td><b>' + esc(data.name || '') + '</b></td></tr>' +
        '<tr><td class="lbl">Father\'s / Husband\'s Name</td><td>' + esc(data.father || '') + '</td></tr>' +
        '<tr><td class="lbl">Date of Birth & Gender</td><td>' + esc(data.dob || '') + ' · ' + esc(data.gender || '—') + '</td></tr>' +
        '<tr><td class="lbl">Marital Status</td><td>' + esc(data.marital || '') + '</td></tr>' +
        '<tr><td class="lbl">Primary Mobile</td><td>' + esc(data.mobile || '') + '</td></tr>' +
        '<tr><td class="lbl">Alternate Contact</td><td>' + esc(data.altmobile || '—') + '</td></tr>' +
        '<tr><td class="lbl">Email Address</td><td>' + esc(data.email || '—') + '</td></tr>' +
        '<tr><td class="lbl">Languages Known</td><td>' + esc(data.languages || '—') + '</td></tr>' +
        '<tr><td class="lbl">Blood Group</td><td>' + esc(data.blood || '—') + '</td></tr>' +
        '<tr><td class="lbl">Present Address</td><td>' + esc(data.addr_present || '') + '</td></tr>' +
        '<tr><td class="lbl">Permanent Address</td><td>' + esc(data.addr_permanent || 'Same as present address') + '</td></tr>' +
        '<tr><td class="lbl">Emergency Contact</td><td>' + esc(data.emg_name || '') + ' (' + esc(data.emg_rel || 'Contact') + ') - ' + esc(data.emg_mobile || '') + '</td></tr>' +
      '</table>' +
    '</div>' +

    // 3. Bank & Identity
    '<div class="card">' +
      '<div class="card-title">3. Bank Account & Identity Documents</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Aadhaar Card Number</td><td><b>' + esc(data.aadhaar || '') + '</b></td></tr>' +
        '<tr><td class="lbl">PAN Card Number</td><td><b>' + esc(data.pan || '—') + '</b></td></tr>' +
        '<tr><td class="lbl">Bank Name & Branch</td><td>' + esc(data.bank || '') + '</td></tr>' +
        '<tr><td class="lbl">Account Number</td><td><b>' + esc(data.account || '') + '</b></td></tr>' +
        '<tr><td class="lbl">IFSC Code</td><td><b>' + esc(data.ifsc || '') + '</b></td></tr>' +
        '<tr><td class="lbl">Documents Submitted</td><td>' + esc(data.docs || 'None checked') + '</td></tr>' +
        aadhaarDocHtml +
      '</table>' +
    '</div>' +

    // 4. Provident Fund & ESIC
    '<div class="card">' +
      '<div class="card-title">4. Provident Fund (EPFO Form 11) & ESIC</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Worked Before?</td><td>' + esc(data.worked || 'No') + (data.prev_employer ? ' (At: ' + esc(data.prev_employer) + ', Last day: ' + esc(data.prev_last_day) + ')' : '') + '</td></tr>' +
        '<tr><td class="lbl">Prior PF Account / UAN</td><td>' + esc(data.epf || 'No') + '</td></tr>' +
        '<tr><td class="lbl">Universal Account Number (UAN)</td><td>' + esc(data.uan || 'None') + '</td></tr>' +
        '<tr><td class="lbl">Previous Member / Establishment ID</td><td>' + esc(data.pf_account || '—') + ' · ' + esc(data.prev_estb || '—') + '</td></tr>' +
        '<tr><td class="lbl">Pension Scheme (EPS 1995)</td><td>' + esc(data.eps || '—') + '</td></tr>' +
        '<tr><td class="lbl">ESIC Insurance Number (IP)</td><td>' + esc(data.esic || 'To be generated') + '</td></tr>' +
        '<tr><td class="lbl">International Worker?</td><td>' + esc(data.intl || 'No') + '</td></tr>' +
      '</table>' +
    '</div>' +

    // 5. Nominee & Family
    '<div class="card">' +
      '<div class="card-title">5. Nominee & Family Dependents</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Primary Nominee</td><td>' + esc(data.nom_name || '') + ' (' + esc(data.nom_rel || '') + ')' + (data.nom_dob ? ' · DOB: ' + esc(data.nom_dob) : '') + '</td></tr>' +
        familyRows +
      '</table>' +
    '</div>' +

    // 6. Declarations & Signature
    '<div class="card">' +
      '<div class="card-title">6. Statutory Undertakings & Digital Signature</div>' +
      '<table class="table">' +
        '<tr><td class="lbl">Declarations Accepted</td><td><b>' + declCount + ' of 5</b> mandatory undertakings confirmed</td></tr>' +
        '<tr><td class="lbl">PF Exemption Request</td><td>' + (data.declarations && data.declarations.dec_pf_exclude ? 'YES — Requested Excluded Employee status (Separate undertaking attached)' : 'No') + '</td></tr>' +
        '<tr><td class="lbl">Signed By</td><td><b>' + esc(data.sig_name || data.name || '') + '</b> at ' + esc(data.sig_place || 'Hyderabad') + '</td></tr>' +
        '<tr><td class="lbl">Signature Timestamp</td><td>' + esc(data.signedAt || todayStr) + '</td></tr>' +
        (data.signatureImage ? '<tr><td class="lbl">Digital Signature</td><td><img class="sig-img" src="' + data.signatureImage + '"/></td></tr>' : '') +
      '</table>' +
    '</div>' +

    '<div class="footer">' +
      'Generated by TrooGood Onboarding Portal · Mformillet Foods Pvt Ltd · Stored in Google Drive' +
    '</div>' +

    '</body></html>';

  var htmlBlob = Utilities.newBlob(html, 'text/html', empName + ' - Joining Form.html');
  var pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName(empName + ' - Joining Form (' + data.reference + ').pdf');

  var file = subFolder.createFile(pdfBlob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return file;
}

/**
 * Builds and saves dedicated "PF Exemption Undertaking & Self-Declaration" PDF
 */
function createPfExemptionPdf(data, subFolder, empName) {
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMMM yyyy, hh:mm a');
  
  var html = '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<style>' +
    'body { font-family: Helvetica, Arial, sans-serif; color: #0f172a; margin: 0; padding: 30px; font-size: 12.5px; line-height: 1.6; background: #ffffff; }' +
    '.header { border-bottom: 3px solid #00B5E8; padding-bottom: 16px; margin-bottom: 20px; }' +
    '.brand { font-size: 14px; font-weight: 800; color: #00B5E8; text-transform: uppercase; letter-spacing: 0.1em; }' +
    'h1 { margin: 6px 0 4px; font-size: 20px; color: #0f172a; }' +
    '.sub { font-size: 12px; color: #475569; font-weight: 600; margin-bottom: 6px; }' +
    '.meta-card { border: 1.5px solid #bcecf8; background: #f0f9ff; border-radius: 6px; padding: 12px 16px; margin-bottom: 22px; }' +
    '.meta-table { width: 100%; border-collapse: collapse; }' +
    '.meta-table td { padding: 4px 8px; font-size: 12px; vertical-align: top; }' +
    '.meta-table td.lbl { width: 32%; font-weight: bold; color: #0369a1; }' +
    '.statute-box { background: #fffbeb; border: 1.5px solid #fde68a; border-left: 5px solid #d97706; padding: 12px 16px; border-radius: 4px; margin-bottom: 22px; font-size: 12px; color: #92400e; }' +
    '.clause-list { padding-left: 20px; margin: 16px 0 24px; }' +
    '.clause-list li { margin-bottom: 12px; font-size: 12.5px; color: #1e293b; text-align: justify; }' +
    '.sig-card { border: 1px solid #cbd5e1; border-radius: 6px; padding: 16px; margin-top: 30px; background: #f8fafc; }' +
    '.sig-img { max-width: 250px; max-height: 80px; display: block; border-bottom: 1.5px solid #0f172a; margin: 10px 0 8px; }' +
    '.footer { margin-top: 40px; padding-top: 12px; border-top: 1px dashed #cbd5e1; font-size: 10px; color: #94a3b8; text-align: center; }' +
    '</style></head><body>' +

    '<div class="header">' +
      '<div class="brand">TROOGOOD · Mformillet Foods Pvt Ltd</div>' +
      '<h1>FORM 11 ANNEXURE: STATUTORY SELF-DECLARATION FOR PF EXEMPTION</h1>' +
      '<div class="sub">Declaration of "Excluded Employee" under Paragraph 2(f) of the Employees\' Provident Funds Scheme, 1952</div>' +
    '</div>' +

    '<div class="meta-card">' +
      '<table class="meta-table">' +
        '<tr><td class="lbl">Reference ID:</td><td><b>' + esc(data.reference || '') + '</b></td>' +
        '<td class="lbl">Date of Joining:</td><td>' + esc(data.doj || todayStr) + '</td></tr>' +
        '<tr><td class="lbl">Employee Name:</td><td><b>' + esc(empName) + '</b></td>' +
        '<td class="lbl">Unit / Location:</td><td>' + esc(data.unit || '') + '</td></tr>' +
        '<tr><td class="lbl">Designation:</td><td>' + esc(data.desig || '') + '</td>' +
        '<td class="lbl">Monthly Gross Salary:</td><td><b>' + (data.salary ? '₹' + esc(data.salary) : 'Exceeds ₹15,000/-') + '</b></td></tr>' +
        '<tr><td class="lbl">Aadhaar Number:</td><td>' + esc(data.aadhaar || '') + '</td>' +
        '<td class="lbl">PAN Number:</td><td>' + esc(data.pan || '—') + '</td></tr>' +
      '</table>' +
    '</div>' +

    '<div class="statute-box">' +
      '<b>Statutory Definition (EPF Scheme 1952, Para 2(f)):</b><br>' +
      'An "Excluded Employee" means an employee whose pay at the time of joining employment exceeds ₹15,000/- per month and who has never previously been a member of the Employees\' Provident Fund or Employees\' Pension Scheme.' +
    '</div>' +

    '<h3 style="margin-bottom:8px; color:#0f172a;">Solemn Undertaking & Self-Declaration:</h3>' +
    '<ol class="clause-list">' +
      '<li>I, <b>' + esc(empName) + '</b>, hereby solemnly declare that I have <b>never been an enrolled member</b> of the Employees\' Provident Fund (EPF) Scheme, 1952 or the Employees\' Pension Scheme (EPS), 1995 with any previous establishment.</li>' +
      '<li>I confirm that I have <b>never been allotted a Universal Account Number (UAN)</b> by the Employees\' Provident Fund Organisation (EPFO) or by any former employer.</li>' +
      '<li>I declare that my monthly wages / gross pay at the time of joining TrooGood (Mformillet Foods Pvt Ltd) <b>exceeds the statutory wage threshold of ₹15,000/- per month</b>.</li>' +
      '<li>In light of the above facts, I hereby voluntarily request the management of TrooGood / Mformillet Foods Pvt Ltd to treat me as an <b>"Excluded Employee"</b> in terms of Paragraph 2(f) of the EPF Scheme, 1952.</li>' +
      '<li>I understand and declare that neither employee PF contribution nor employer PF contribution shall be deducted or remitted on my behalf.</li>' +
      '<li>I affirm that this self-declaration is made <b>voluntarily of my own free will</b>, with complete understanding of the statutory rules, and without any coercion or inducement whatsoever.</li>' +
    '</ol>' +

    '<div class="sig-card">' +
      '<b>Digital Signature & Candidate Confirmation:</b><br>' +
      (data.signatureImage ? '<img class="sig-img" src="' + data.signatureImage + '" />' : '<div style="height:40px;">[Signed Electronically]</div>') +
      '<div style="font-size:12px; color:#334155; margin-top:4px;">' +
        'Candidate Name: <b>' + esc(data.sig_name || empName) + '</b><br>' +
        'Place / Location: <b>' + esc(data.sig_place || 'Hyderabad') + '</b><br>' +
        'Submission Timestamp: <b>' + esc(data.signedAt || todayStr) + '</b>' +
      '</div>' +
    '</div>' +

    '<div class="footer">' +
      'TrooGood Statutory HR Compliance Record · Retained permanently in Google Drive folder: ' + esc(data.reference) +
    '</div>' +

    '</body></html>';

  var htmlBlob = Utilities.newBlob(html, 'text/html', empName + ' - PF Exemption Declaration.html');
  var pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName(empName + ' - PF Exemption Declaration (' + data.reference + ').pdf');

  var file = subFolder.createFile(pdfBlob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return file;
}

/**
 * Optional email alert to HR
 */
function notify(data, folderUrl, pdfUrl, pfPdfUrl) {
  if (!HR_EMAIL) return;
  try {
    var body = 'A new onboarding joining form has been submitted and saved in Google Drive.\n\n' +
      'Reference ID: ' + (data.reference || '') + '\n' +
      'Employee Name: ' + (data.name || '') + '\n' +
      'Unit: ' + (data.unit || '') + '\n' +
      'Role: ' + (data.desig || '') + '\n' +
      'Salary: ' + (data.salary ? '₹' + data.salary : 'N/A') + '\n' +
      'Aadhaar: ' + (data.aadhaar || 'N/A') + '\n' +
      'Mobile: ' + (data.mobile || 'N/A') + '\n' +
      'PF Exemption Requested: ' + (pfPdfUrl ? 'YES (Undertaking PDF generated)' : 'No') + '\n\n' +
      'Drive Folder: ' + folderUrl + '\n' +
      'Joining Form PDF: ' + pdfUrl +
      (pfPdfUrl ? '\nPF Exemption PDF: ' + pfPdfUrl : '');

    MailApp.sendEmail(
      HR_EMAIL,
      'New Joining Form: ' + (data.reference || '') + ' - ' + (data.name || '') + ' (' + (data.unit || '') + ')',
      body
    );
  } catch(e) {}
}


function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function esc(str) {
  return String(str || '').replace(/[&<>"]/g, function(m) {
    return { '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;' }[m];
  });
}
