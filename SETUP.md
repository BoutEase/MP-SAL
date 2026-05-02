# MP-SAL Payroll — Setup Guide

Complete setup takes about 20 minutes. No coding required.

---

## What you will set up

```
Petpooja Excel  →  index.html (Netlify)  →  Google Apps Script  →  Google Sheets
     ↑                  ↑                          ↑                     ↑
  You upload        Your web app              Your backend           Your database
```

---

## Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and click **+ Blank**.
2. Name it **MP-SAL Payroll** (top-left, click "Untitled spreadsheet").
3. Look at the URL in your browser. It looks like:
   ```
   https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXX/edit
   ```
   Copy the long string between `/d/` and `/edit`. This is your **Spreadsheet ID**.
   
   > Example: `1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms`
   
4. Keep this tab open.

---

## Step 2 — Create the Google Apps Script project

1. Go to [script.google.com](https://script.google.com) and click **+ New project**.
2. Name it **MP-SAL Backend** (click "Untitled project" at the top).
3. You will see a file called `Code.gs` with an empty function. **Delete all of it.**
4. Open the `Code.gs` file from this repository and **copy the entire contents**.
5. Paste it into the Apps Script editor.
6. Press **Ctrl+S** (or Cmd+S on Mac) to save.

---

## Step 3 — Add your Spreadsheet ID to Apps Script

1. In the Apps Script editor, click the **gear icon (⚙)** on the left sidebar — this is **Project Settings**.
2. Scroll down to **Script Properties**.
3. Click **Add script property**.
4. Set:
   - **Property**: `SPREADSHEET_ID`
   - **Value**: paste the ID you copied in Step 1
5. Click **Save script properties**.

---

## Step 4 — Deploy the Apps Script as a Web App

1. In the Apps Script editor, click **Deploy** (top-right) → **New deployment**.
2. Click the **gear icon** next to "Select type" → choose **Web app**.
3. Fill in the settings:
   - **Description**: `MP-SAL v1`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Click **Deploy**.
5. Google will ask you to **authorise** the app:
   - Click **Authorise access**
   - Choose your Google account
   - You may see "Google hasn't verified this app" — click **Advanced** → **Go to MP-SAL Backend (unsafe)**
   - Click **Allow**
6. After authorisation, you will see a **Web app URL** that looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```
   **Copy this URL.** You will need it in Step 6.

> **Important:** Every time you edit `Code.gs` later, you must deploy a **new version**:
> Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy.

---

## Step 5 — Deploy the frontend to Netlify

### Option A — Drag and drop (easiest, no account needed for basic use)

1. Go to [netlify.com](https://netlify.com) and sign up for a free account.
2. From your Netlify dashboard, scroll down to the **"Deploy manually"** section.
3. Drag and drop a folder containing these four files:
   - `index.html`
   - `app.js`
   - `manifest.json`
   - `icon.svg`
4. Netlify will give you a URL like `https://random-name-123.netlify.app`.

### Option B — Deploy from GitHub (recommended, auto-updates when you push)

1. Go to [netlify.com](https://netlify.com) → **Add new site** → **Import an existing project**.
2. Connect to GitHub and select the **MP-SAL** repository.
3. Set:
   - **Branch**: `claude/payroll-web-app-kSucH`
   - **Build command**: *(leave empty)*
   - **Publish directory**: `/` (or leave as `.`)
4. Click **Deploy site**.
5. Netlify will give you a URL. You can rename it under **Site settings → Site name**.

---

## Step 6 — First-time app configuration

1. Open your Netlify URL in a browser (works on mobile too).
2. You will see a **"Connect to Google Sheets"** screen.
3. Paste the **Web App URL** from Step 4 into the box.
4. Click **Connect & Initialise Sheets**.
5. If successful, you will see "Connected!" and the app will create all 7 tabs in your Google Sheet automatically:
   - Employees, Holidays, Payroll, Advances, Bonus, Payments, Settings

6. You will be taken to the **Login screen**.

---

## Step 7 — First login and configuration

### Login as Admin
- Select **Admin** tab
- Default PIN: **1234**
- Click **Login**

### Change your PINs immediately
1. Go to the **Settings** tab (top navigation)
2. Under **General**, change:
   - **Admin PIN** to something only you know
   - **Manager PIN** to something your manager knows
3. Click **Save All Settings**

### Add your holidays
1. Still in **Settings**, scroll to **Holidays**
2. Add your 8 company holidays for the year (date + name)
   > You can change these any time — useful when Eid or other lunar dates shift
3. Click **Save All Settings**

### Add your employees
1. Go to the **Employees** tab
2. Click **+ Add Employee** for each employee
3. Fill in:
   - **Name**: Use the exact same spelling you want on payslips
   - **Team**: Emb or Stitching
   - **Petpooja ID**: The Employee ID number from Petpooja (e.g., 2, 3, 4…)
   - **Petpooja Name**: The name as it appears in Petpooja export (if different from above)
   - **Weekly Salary**: Their weekly salary in ₹
4. Click **Save** for each employee

> **Tip on Petpooja ID:** Open a Petpooja export Excel, find an employee's block header. It shows `Employee ID: 2`. That number goes in **Petpooja ID**. This is how the app auto-matches employees when you upload the Excel each month.

---

## Step 8 — Running your first payroll

1. Go to the **Payroll** tab.
2. Select the correct **month** from the month picker.
3. Click **Upload Petpooja Excel** and select the attendance export file.
4. The app will:
   - Parse every employee block
   - Match each to your employee database by Petpooja ID
   - Calculate salary, OT, shortfall deductions, advances, and bonus cuts automatically
5. If any employees are **not matched** (shown in orange), use the dropdown to manually map them to the correct employee, then click **Apply**.
6. Review the payroll table. Click **Edit** on any row to adjust:
   - Full Days, Half Days, OT minutes, shortfall minutes
   - Advances (if different from what was recorded)
   - Any notes
   - Click **Recalculate & Apply** after editing
7. Click **Save Draft** to save your work in progress.
8. When everything is correct, click **Finalise Month**.
   - This locks all records
   - Automatically updates the Bonus Pool for each eligible employee
   - Records cannot be edited after finalisation

---

## Step 9 — Manager advance submission

1. Share your Netlify URL with your manager.
2. On the login screen, they select **Manager** and enter the Manager PIN.
3. They can only see the **Submit Advance** page:
   - Select employee
   - Enter date, amount, mode (Cash/GPay), and which month to deduct from
   - Submit
4. All advances they submit appear instantly in your **Advances** tab (Admin view) and are automatically included in the next payroll calculation for that employee.

---

## How the salary calculation works (quick reference)

| Item | Rule |
|---|---|
| Daily rate | Weekly salary ÷ 6 |
| Hourly rate | Daily rate ÷ 9 |
| Paid days (TD) | Full Days + Half Days + Paid holidays where absent |
| Gross pay | TD × Daily rate |
| Weekday OT | Only if total hours ≥ 10h. OT = (hours − 9h) rounded down to nearest 15 min × hourly rate |
| Sunday OT | All hours worked on Sunday (+ 1h lunch if enabled) × hourly rate |
| Holiday OT | All hours worked on paid holiday (+ 1h lunch if enabled) × hourly rate |
| Half-day shortfall | (9h − actual hours that day) × hourly rate, deducted |
| Shortfall on FD | Only deducted if total hours < 8.5h (configurable in Settings) |
| Bonus cut | ₹500/month deducted if employee worked ≥ 6 days |
| Advance deduction | Sum of all advances recorded for that month |
| **Net pay** | Gross + OT − Shortfall − Bonus cut − Advances |

---

## Bonus Pool explained

- Every month an eligible employee (worked ≥ 6 days) gets ₹500 deducted from salary.
- The company adds ₹500 → total ₹1,000 added to their pool.
- When the pool reaches ₹24,000 → it pays out automatically when you finalise that month's payroll → pool resets to zero.
- You can see every employee's current balance under the **Bonus** tab.

---

## Making updates to Code.gs later

Whenever the `Code.gs` file is updated in this repository:

1. Open [script.google.com](https://script.google.com) → your MP-SAL Backend project.
2. Replace the content of `Code.gs` with the new version.
3. Save (Ctrl+S).
4. Click **Deploy** → **Manage deployments**.
5. Click the pencil (Edit) on your existing deployment.
6. Change **Version** to **New version**.
7. Click **Deploy**.

The frontend (`index.html`, `app.js`) updates automatically on Netlify if you connected GitHub.

---

## Troubleshooting

**"Unauthorised" error on login**
- Your PIN is wrong, or the GAS URL is incorrect. Click **Change URL** on the login screen and re-enter it.

**"Cannot open spreadsheet" error during setup**
- Make sure the Spreadsheet ID in Script Properties is correct (no extra spaces).
- Make sure the Google account that owns the Apps Script is the same one that owns the Sheet.

**Employees not matching after Petpooja upload**
- The Petpooja ID in your Employees tab doesn't match what's in the Petpooja export.
- Open the Petpooja Excel, find the employee's block header (`Employee ID: X`), and update the Petpooja ID in the Employees tab.

**Changes to Code.gs not taking effect**
- You edited the script but forgot to deploy a new version. See "Making updates" above.

**App not loading on mobile**
- Make sure your Netlify URL uses `https://`. The PWA install prompt only works on HTTPS.

**"Add to Home Screen" on iPhone**
- Open the app in Safari → tap the Share button → scroll down → tap **Add to Home Screen**.

**"Add to Home Screen" on Android**
- Open the app in Chrome → tap the three-dot menu → tap **Add to Home screen** or **Install app**.

---

## Default settings reference

| Setting | Default | Where to change |
|---|---|---|
| Admin PIN | 1234 | Settings → General |
| Manager PIN | 5678 | Settings → General |
| Shortfall threshold | 8.5 hours | Settings → Calculation Rules |
| OT minimum | 10 hours | Settings → Calculation Rules |
| OT rounding | 15 minutes | Settings → Calculation Rules |
| Add lunch to Sunday/Holiday OT | Yes | Settings → Calculation Rules |
| Lunch duration | 1 hour | Settings → Calculation Rules |
| Bonus pool target | ₹24,000 | Settings → Bonus Pool |
| Employee contribution | ₹500/month | Settings → Bonus Pool |
| Company contribution | ₹500/month | Settings → Bonus Pool |
| Bonus eligibility | 6 days minimum | Settings → Bonus Pool |
