# 📋 Leave Request System
[![Codacy Badge](https://app.codacy.com/project/badge/Grade/9d847cad0c874e0f9ec4e3948080117a)](https://app.codacy.com/gh/anugrahiyyan/Approval-Leave-Request/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

A fully automated **Google Apps Script-based Leave Request System** with multi-stage approvals.

---

## ✨ Features

- 📝 Google Form frontend with enhanced UX
- 📆 Integration with Google Calendar (Add to calendar for reporting team)
- 🔄 Multi-stage approval workflow: **SPV → HR → GM (optional)**
- 📧 Dynamic email notifications with styled Bootstrap-based HTML templates
- 📨 Final decision notifications sent to requester and reporting team
- 📊 Google Sheet-powered backend for easy tracking and management

## ✨ New Feature
- ✅ Unique one-time tokens per approval link
- ❌ Prevents multiple approvals or external tampering
- 🚫 Displays a styled error message if token is missing, invalid, or already used

---

## 🚀 Tech Stack

- [Google Apps Script](https://developers.google.com/apps-script)
- Google Sheets
- Google Workspace (Gmail, Calendar)
- HTML + Bootstrap (Email + UI Styling)
- JavaScript (Client-side interaction)

---

## 🚀 How to Run This Project

Follow the steps below to set everything up:

1. **Create a new Google Sheet**.
2. **Rename the first sheet** to `Requests`  
   > This sheet will serve as the backend for processing requests.
3. **(Optional)** Create additional sheets for dashboards or analytics — feel free to customize them to fit your needs.
4. **Open the Apps Script editor**  
   Go to `Extensions` → `Apps Script`.
5. **Add all script files except `index.html`** to your Apps Script project.
6. **Deploy your project** using the Apps Script deployment options.

---

## ❓ Why is `index.html` not included in Apps Script?

- The `index.html` file is used **only for the landing page**, hosted separately.
- This approach helps ensure proper **domain ownership verification** for Google Cloud Platform (GCP).
- Once verified, your GCP app can be associated securely with your custom domain.
- You can read more about GCP domain verification via [Google’s documentation](https://support.google.com/cloud/answer/9110914).

---
