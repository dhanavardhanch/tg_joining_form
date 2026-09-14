# TrooGood Employee Onboarding & Joining Portal

> **Official Digital Onboarding Portal for TrooGood · Mformillet Foods Pvt Ltd**

A fast, mobile-friendly multi-step web application designed to onboard new factory and corporate employees seamlessly. Submissions are saved directly into designated Google Drive employee folders and logged to Google Sheets with zero manual data entry.

---

## 🌟 Key Features

* **TrooGood Brand Identity**: Built in official **TrooGood Cyan (`#00B5E8`)** and crisp whites (`#FFFFFF`), with zero green colors and high-contrast typography.
* **100% Mobile Responsive**: Optimized for factory floor and field onboarding across Android and iPhone browsers with touch-friendly input targets (>= 48px).
* **Employee Passport Photo Upload**:
  * Take a live camera photo or select from gallery.
  * Automatic client-side canvas compression down to ~50KB to ensure fast submission without lag.
  * Embedded at the **top-left** of the generated application form PDF.
* **Aadhaar Card Document / Photo Upload**:
  * Dedicated document capture box alongside the 12-digit Aadhaar number field.
  * Compressed and stored directly in the candidate's Drive subfolder.
* **Sequential Reference ID**:
  * Auto-generates clean reference IDs starting at `TG-JF-001` (e.g. `TG-JF-001`, `TG-JF-002`, etc.).
* **Capital First-Letter Enforcement**:
  * Text fields automatically format typed input into Title Case (`dhana` → `Dhana`).
  * Real-time validation warning if text does not begin with a capital letter.
* **Disabled Employee Code**:
  * Empty and disabled by default with an *"Auto-allotted by HR"* badge.
* **Touch-Enabled Digital Signature**:
  * Smooth finger/stylus canvas drawing with clear and *"Adopt Typed Name as Signature"* option.
* **Statutory Compliance (EPFO Form 11 & ESIC)**:
  * Collects all mandatory provident fund history, UAN, EPS 1995 particulars, bank account details, IFSC, and nominee details.

---

## 📁 Google Drive & Sheet Automation

Each submission executes via `TrooGood_Form_To_Google_Sheet.gs` to automatically organize HR files:

```text
📁 Target Google Drive Folder: 1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t
   └── 📁 [Candidate Name] - [Reference ID]
        ├── 📄 [Candidate Name] - Joining Form ([Reference ID]).pdf  (with top-left photo)
        ├── 🖼️ Photograph - [Candidate Name].jpg
        ├── 🪪 Aadhaar - [Candidate Name].jpg
        └── 🖊️ Signature - [Candidate Name].png
```

### Google Sheet Columns Logged:
1. `Timestamp`
2. `Reference ID`
3. `Candidate Name`
4. `Aadhaar Number`
5. `PAN Number`
6. `Bank Account Number`
7. `Bank IFSC Code`
8. `ESIC Number`
9. `Unit / Location`
10. `Role / Designation`
11. `Mobile Number`
12. `Date of Joining`
13. `Drive Subfolder Link`
14. `PDF Application Link`
15. `Passport Photo Link`
16. `Aadhaar Card Link`

---

## 🚀 Deployment Options

### Option 1: Vercel (Recommended)
1. Import this repository into [Vercel](https://vercel.com/new).
2. Deploy directly with default static settings.
3. Or deploy via terminal:
   ```bash
   vercel --prod
   ```

### Option 2: Netlify Drop (Zero CLI)
1. Go to [app.netlify.com/drop](https://app.netlify.com/drop).
2. Drag and drop this project folder into the browser window.
3. Instant live HTTPS URL generated in 10 seconds.

### Option 3: Local Testing
Run a local web server in the project folder:
```bash
python -m http.server 8080
```
Open `http://localhost:8080` in your browser.

---

## ⚙️ Google Apps Script Setup

1. Open your target Google Sheet or go to [script.google.com](https://script.google.com).
2. Copy the contents of [`TrooGood_Form_To_Google_Sheet.gs`](./TrooGood_Form_To_Google_Sheet.gs) into the script editor.
3. Verify your target Google Drive folder ID in line 18:
   ```javascript
   var PARENT_FOLDER_ID = '1RPOKrjrlykX2GjkbQ0wwBFqiIXV8zs4t';
   ```
4. Click **Deploy** → **New deployment**:
   * Type: **Web app**
   * Execute as: **Me**
   * Who has access: **Anyone**
5. Copy the Web App URL (ends with `/exec`).
6. Update `SUBMIT_URL` in `TrooGood_Joining_Form.html` (line ~1775).

---

## 📄 License
Internal HR tool developed for **TrooGood · Mformillet Foods Pvt Ltd**. All rights reserved.
