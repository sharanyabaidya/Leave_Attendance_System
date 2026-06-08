# Leave & Attendance Management System

This repository contains the refactored Object-Oriented backend and dynamic front-end layout for the **GYWS Leave & Attendance Management System**. 

The system leverages Google Apps Script as a Serverless Application engine, with Google Sheets functioning as the primary relational database layer.

---

## 🚀 Step-by-Step Deployment Guide

### Step 1: Create a Google Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.
2. Name it (e.g. `GYWS Leave & Attendance Database`).
3. You do **not** need to create the individual sheets manually. The setup script in Step 3 will automatically format and create them for you.

### Step 2: Open the Apps Script Editor
1. In your new Google Spreadsheet, go to the top menu and select **Extensions** > **Apps Script**.
2. This opens the container-bound script editor.
3. Create the following files in the editor and copy-paste the corresponding file contents from this repository:
   * **`Code.gs`** (Script)
   * **`OOP.gs`** (Script)
   * **`Index.html`** (HTML)
   * **`AppScript.html`** (HTML)
   * **`Styles.html`** (HTML)

### Step 3: Initialize Database (Run Dummy Data Setup)
To quickly pre-fill your sheets with correct headers, formatting, and sample user profiles (with hashed passwords):
1. In the Apps Script editor toolbar, look at the function selector dropdown.
2. Select **`setupDummyData`**.
3. Click the **Run** button.
4. When prompted, click **Review Permissions** and authorize the script to access your spreadsheet and emails (required for notifications).
5. Open your Google Spreadsheet. You will see 7 sheets created and populated with sample rows:
   * `Employees` (Includes employee, HR, office staff, and admin logins)
   * `LeaveApplications` (Pre-filled leave entries)
   * `Attendance` (Pre-filled attendance summaries)
   * `LeaveBalance` (Synchronized balances)
   * `Config` (Global settings)
   * `AuditLog` (Action audits)
   * `SalaryCalculations` (Payroll reports)

### Step 4: Deploy as a Web App
To host the user interface live:
1. At the top right of the Apps Script editor, click **Deploy** > **New deployment**.
2. Click the gear icon next to "Select type" and select **Web app**.
3. Configure the settings exactly as follows:
   * **Description**: `Version 1.0.0`
   * **Execute as**: `Me (your-email@gmail.com)`
   * **Who has access**: `Anyone` *(Crucial: This allows your employees to access the login form without requesting sheet permission)*.
4. Click **Deploy**.
5. Copy the **Web App URL** (e.g., `https://script.google.com/macros/s/.../exec`). This is the live URL your team will use to log in and submit leaves.

---

## 🔑 Default Credentials (from Dummy Data)
Use these logins to test the interface after deployment:
* **Employee**: ID: `EMP001` | Password: `pass123`
* **HR Manager**: ID: `EMP002` | Password: `hr123`
* **Office Staff**: ID: `EMP003` | Password: `pass123`
* **Super Admin**: ID: `EMP004` | Password: `admin123`
