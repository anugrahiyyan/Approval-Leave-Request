## 🗂 ONEderland Leave Request System — CHANGELOG

---

### 🛠️ Initial Version (May 2025)

**Core Functionality:**

* Custom HTML form (`form.html`) using `doGet()`.
* Form data saved to `Requests` sheet.
* Email approval workflow: SPV → HR → GM → Reporting.
* Token-based links for approval in email.

---

### 📬 Approval Workflow & Email Styling (May 2025)

* Bootstrap email templates (`emailtemplate.html`, `result.html`, `finalNotification.html`).
* Conditional decision flow based on requester role (SPV/HR/GM).
* Rejection emails now include full decision history.

---

### 🎆 UX Enhancements (Late May 2025)

* Fireworks animation and modal on success.
* Button changes to "Success" and resets form.
* Enhanced timestamp formatting: `dd-mm-yyyy hh:mm:ss AM/PM`.

---

### 🔐 Token-Based Approval Logic (June 2025)

* Each stage (SPV/HR/GM) receives unique 8-character token.
* Prevents re-approving or invalid access.
* Token is marked as "used" after action.

---

### 🧮 Leave Balance Tracking (July 2025)

* Balance columns (W: email, X: leave, Y: sick leave).
* Automatically deducts balance on final approval.
* Displays current balance to requester in form success and final email.
* Uses dynamic lookup (no hardcoded ranges).

---

### 🚫 Auto-Rejection Logic (July 2025)

* Checks if requested days exceed balance.
* Automatically rejects with note: `Auto-rejected: Only X day(s) left, but Y requested.`
* Reason added to sheet and final email.

---

### 📩 Final Notification Overhaul (July 22, 2025)

* Enhanced styling for rejection and approval.
* Clear explanation if system overrode approval.
* Added balance summary and advice for next steps.

---

### ✅ Current Feature Summary

* Multi-stage approval with role-based flow.
* Token-secured approval links.
* Balance validation and auto-rejection.
* Fully styled HTML notifications.
* Leave balance updates in real-time.
* User-friendly modals and alerts.

---
