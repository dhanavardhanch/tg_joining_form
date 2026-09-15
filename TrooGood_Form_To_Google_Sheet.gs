/**
 * TrooGood Employee Onboarding → Google Drive PDF & Subfolder Saver
 * ================================================================
 * Saves every onboarding form submission directly into your designated Google Drive folder:
 * Target Folder: https://drive.google.com/drive/folders/1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t
 *
 * For each submission, this script:
 *  1. Creates a subfolder inside the target folder named with the Employee's Name & Reference ID.
 *  2. Saves the employee's uploaded Photograph (.jpg) inside that subfolder.
 *  3. Saves the employee's uploaded Aadhaar Card document (.jpg) inside that subfolder.
 *  4. Formats all details into a clean, professional PDF document with the passport photo at the top-left,
 *     styled in official TrooGood Cyan (#00B5E8) & White, and saves it in that subfolder.
 *  5. Saves the employee's digital signature image as a PNG in the same subfolder.
 *  6. Logs all employee details including Aadhaar, PAN, Bank Account, IFSC, and ESIC into the Google Sheet.
 *  7. (Optional) Emails HR an alert notification with the Drive folder link.
 */

// Target Google Drive Folder ID from user link
var PARENT_FOLDER_ID = '1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t';
var HR_EMAIL         = '';              // e.g. 'hr@troogood.in' — leave blank for no email alerts
var SHEET_NAME       = 'Submissions';   // Sheet tab name for backup logging

/** Web App POST endpoint */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    
    // 1. Locate the parent Google Drive folder
    var parentFolder;
    try {
      parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    } catch (fErr) {
      var folders = DriveApp.getFoldersByName("TrooGood Employee Onboarding Documents");
      parentFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder("TrooGood Employee Onboarding Documents");
    }

    // 2. Locate existing subfolder or create new subfolder for employee
    var empName = (data.name || 'New Employee').trim();
    var refId = (data.reference || '').trim();
    var folderTitle = empName + (data.code ? ' (' + data.code + ')' : '') + (refId ? ' - ' + refId : '');
    
    var subFolder = null;
    var existingFolders = parentFolder.getFolders();
    while (existingFolders.hasNext()) {
      var f = existingFolders.next();
      var fName = f.getName();
      if ((refId && fName.indexOf(refId) > -1) || (empName && fName.toLowerCase().indexOf(empName.toLowerCase()) > -1)) {
        subFolder = f;
        break;
      }
    }
    if (!subFolder) {
      subFolder = parentFolder.createFolder(folderTitle);
      try {
        subFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) {}
    }

    // 3. Save uploaded photograph into the subfolder
    var photoUrl = '';
    if (data.photoImage && data.photoImage.indexOf('base64,') > -1) {
      try {
        var photoBytes = Utilities.base64Decode(data.photoImage.split('base64,')[1]);
        var photoFile = subFolder.createFile(Utilities.newBlob(photoBytes, 'image/jpeg', 'Photograph - ' + empName + '.jpg'));
        try { photoFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        photoUrl = photoFile.getUrl();
      } catch (err) {}
    }
    data.photoLink = photoUrl;

    // 4. Save uploaded Aadhaar Card document into the subfolder
    var aadhaarUrl = '';
    if (data.aadhaarImage && data.aadhaarImage.indexOf('base64,') > -1) {
      try {
        var aadhaarBytes = Utilities.base64Decode(data.aadhaarImage.split('base64,')[1]);
        var aadhaarFile = subFolder.createFile(Utilities.newBlob(aadhaarBytes, 'image/jpeg', 'Aadhaar - ' + empName + '.jpg'));
        try { aadhaarFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        aadhaarUrl = aadhaarFile.getUrl();
      } catch (err) {}
    }
    data.aadhaarLink = aadhaarUrl;

    // 5. Save signature image into the subfolder
    var sigUrl = '';
    if (data.signatureImage && data.signatureImage.indexOf('base64,') > -1) {
      try {
        var bytes = Utilities.base64Decode(data.signatureImage.split('base64,')[1]);
        var sigFile = subFolder.createFile(Utilities.newBlob(bytes, 'image/png', 'Signature - ' + empName + '.png'));
        try { sigFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
        sigUrl = sigFile.getUrl();
      } catch (err) {}
    }
    data.signatureLink = sigUrl;

    // 6. Generate and save the complete formatted PDF (with top-left photo and TrooGood branding)
    var pdfFile = createEmployeePdf(data, subFolder, empName);
    var pdfUrl = pdfFile ? pdfFile.getUrl() : '';

    // 7. Log all structured identity, bank, and drive details to the active Google Sheet
    try {
      logToSheet(data, subFolder.getUrl(), pdfUrl, photoUrl, aadhaarUrl);
    } catch(sheetErr) {}

    // 8. Optional: notify HR
    notify(data, subFolder.getUrl(), pdfUrl);

    return json({
      ok: true,
      reference: data.reference,
      folderUrl: subFolder.getUrl(),
      pdfUrl: pdfUrl,
      photoUrl: photoUrl,
      aadhaarUrl: aadhaarUrl
    });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** Web App GET endpoint (health check) */
function doGet() {
  return json({
    ok: true,
    message: 'TrooGood Onboarding Drive PDF endpoint is live. Brand color: #00B5E8.'
  });
}

/** Builds and saves a formatted PDF of the employee onboarding form */
function createEmployeePdf(data, subFolder, empName) {
  var todayStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd MMMM yyyy, hh:mm a');
  
  // Format Family Members
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

  // Format Declarations
  var declCount = 0;
  if (data.declarations) {
    for (var k in data.declarations) {
      if (data.declarations[k]) declCount++;
    }
  }

  // Photo element for the top-left of the application form PDF
  var photoHtml = '';
  if (data.photoImage && data.photoImage.indexOf('base64,') > -1) {
    photoHtml = '<img src="' + data.photoImage + '" style="width:96px; height:120px; object-fit:cover; border:2px solid #00B5E8; border-radius:4px; display:block;" alt="Photo" />';
  } else {
    photoHtml = '<div style="width:96px; height:120px; border:1.5px dashed #94a3b8; border-radius:4px; text-align:center; font-size:10px; color:#64748b; background:#f8fafc; padding-top:45px; box-sizing:border-box;">Affix Photo</div>';
  }

  // Aadhaar document block if uploaded
  var aadhaarDocHtml = '';
  if (data.aadhaarImage && data.aadhaarImage.indexOf('base64,') > -1) {
    aadhaarDocHtml = '<tr><td class="lbl">Aadhaar Card Document</td><td>' +
      '<img src="' + data.aadhaarImage + '" style="max-width:220px; max-height:130px; object-fit:contain; border:1.5px solid #00B5E8; border-radius:4px; display:block; margin-top:4px;" />' +
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

    // Application Form Header with Photograph at Top-Left
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
        '<tr><td class="lbl">Date of Birth</td><td>' + esc(data.dob || '') + '</td></tr>' +
        '<tr><td class="lbl">Gender & Marital Status</td><td>' + esc(data.gender || '') + ' · ' + esc(data.marital || '') + '</td></tr>' +
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
        '<tr><td class="lbl">PF Exemption Request</td><td>' + (data.declarations && data.declarations.dec_pf_exclude ? 'YES — Requested Excluded Employee status' : 'No') + '</td></tr>' +
        '<tr><td class="lbl">Signed By</td><td><b>' + esc(data.sig_name || data.name || '') + '</b> at ' + esc(data.sig_place || 'Hyderabad') + '</td></tr>' +
        '<tr><td class="lbl">Signature Timestamp</td><td>' + esc(data.signedAt || todayStr) + '</td></tr>' +
        (data.signatureImage ? '<tr><td class="lbl">Digital Signature</td><td><img class="sig-img" src="' + data.signatureImage + '"/></td></tr>' : '') +
      '</table>' +
    '</div>' +

    '<div class="footer">' +
      'Generated by TrooGood Onboarding Portal · Mformillet Foods Pvt Ltd · Stored in Google Drive' +
    '</div>' +

    '</body></html>';

  // Convert HTML to PDF using Drive Utilities
  var htmlBlob = Utilities.newBlob(html, 'text/html', empName + ' - Joining Form.html');
  var pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName(empName + ' - Joining Form (' + (data.reference || '') + ').pdf');

  var file = subFolder.createFile(pdfBlob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return file;
}

/** Backup logging to sheet with explicit columns for Aadhaar, PAN, Bank & ESIC */
function logToSheet(data, folderUrl, pdfUrl, photoUrl, aadhaarUrl) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  
  var headers = [
    'Timestamp',               // A (1)
    'Reference ID',            // B (2)
    'Employee Code',           // C (3)
    'Name',                    // D (4)
    'Unit',                    // E (5)
    'Designation',             // F (6)
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
    'Contractor Name'          // W (23)
  ];

  // Automatically update Row 1 headers without touching or deleting any existing candidate data
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#00B5E8')
    .setFontColor('#ffffff');

  // Helper to preserve leading zeros in numeric string fields (Aadhaar, Account, Mobile, ESIC)
  function textCell(val) {
    if (val === null || val === undefined) return '';
    var str = val.toString().trim();
    if (!str) return '';
    return str.indexOf("'") === 0 ? str : "'" + str;
  }

  var newRow = [
    new Date(),
    data.reference || '',
    data.code || '',
    data.name || '',
    data.unit || '',
    data.desig || '',
    textCell(data.mobile),
    data.doj || '',
    textCell(data.aadhaar),
    data.pan || '',
    data.uan || 'N/A',
    textCell(data.esic || 'N/A'),
    data.health_ins_no || 'N/A',
    textCell(data.account),
    data.ifsc || '',
    folderUrl || '',
    pdfUrl || '',
    photoUrl || '',
    aadhaarUrl || '',
    data.dob || '',
    data.shift || 'General Shift',
    data.engagement || 'On TrooGood Rolls',
    data.vendor || 'N/A'
  ];

  var lastRow = sheet.getLastRow();
  var existingRowIndex = -1;

  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var cleanAadhaar = (data.aadhaar || '').toString().replace(/\s+/g, '');
    var cleanRef = (data.reference || '').toString().trim();
    var cleanMobile = (data.mobile || '').toString().trim();
    var cleanName = (data.name || '').toString().trim().toLowerCase();

    for (var i = 0; i < dataRange.length; i++) {
      var row = dataRange[i];
      var rowRef = (row[1] || '').toString().trim();
      var rowName = (row[3] || '').toString().trim().toLowerCase();
      var rowAadhaar = (row[8] || '').toString().replace(/\s+/g, '');
      var rowMobile = (row[6] || '').toString().trim();

      var matchAadhaar = (cleanAadhaar.length >= 10 && cleanAadhaar === rowAadhaar);
      var matchRef = (cleanRef.length > 0 && cleanRef === rowRef);
      var matchMobileName = (cleanMobile.length >= 10 && cleanMobile === rowMobile && cleanName === rowName);

      if (matchAadhaar || matchRef || matchMobileName) {
        existingRowIndex = i + 2;
        break;
      }
    }
  }

  if (existingRowIndex > 1) {
    // Candidate already exists -> UPDATE existing row in place instead of creating duplicate
    sheet.getRange(existingRowIndex, 1, 1, newRow.length).setValues([newRow]);
  } else {
    // Candidate is new -> APPEND new row
    sheet.appendRow(newRow);
  }
}

/** Optional email alert */
function notify(data, folderUrl, pdfUrl) {
  if (!HR_EMAIL) return;
  try {
    MailApp.sendEmail(
      HR_EMAIL,
      'New Joining Form: ' + (data.name || '') + ' (' + (data.unit || '') + ')',
      'A new onboarding joining form has been submitted and saved in Google Drive.\n\n' +
      'Employee Name: ' + (data.name || '') + '\n' +
      'Unit: ' + (data.unit || '') + '\n' +
      'Role: ' + (data.desig || '') + '\n' +
      'Reference ID: ' + (data.reference || '') + '\n' +
      'Aadhaar: ' + (data.aadhaar || 'N/A') + '\n' +
      'PAN: ' + (data.pan || 'N/A') + '\n' +
      'Bank Account: ' + (data.account || 'N/A') + '\n\n' +
      'Drive Folder: ' + folderUrl + '\n' +
      'Download PDF: ' + pdfUrl
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
