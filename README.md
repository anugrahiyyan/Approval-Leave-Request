# 📋 ONEderland Leave Request System

A fully automated **Google Apps Script-based Leave Request System** with multi-stage approvals, real-time notifications, and a modern user interface.

---

## ✨ Features

### Core Functionality
- 📝 Modern Google Form frontend with responsive design
- 📆 Automatic Google Calendar integration for approved leaves
- 🔄 Multi-stage approval workflow: **SPV → HR → GM (conditional)**
- 📧 Beautiful Bootstrap-styled HTML email templates
- 📊 Google Sheet-powered backend for tracking and management

### Leave Types Supported
- Annual Leave (Full-day & Half-day)
- Sick Leave (with Medical Certificate upload)
- Bereavement Leave
- Marriage Leave
- Maternity Leave
- Working From Home (WFH)
- Unpaid Leave

### Advanced Features
- ✅ **One-time Token Security** — Prevents duplicate actions or link tampering
- 📋 **Multi-Leave Submission** — Submit multiple leave types in one request
- 🕐 **Half-Day Support** — Clear "0.5 (Half-Day)" display across all views
- 📊 **Approver Dashboard** — View and action pending requests with user avatars
- ❌ **Reject with Notes** — Add rejection reasons for clarity
- 📅 **Automatic Calendar Events** — Approved leaves auto-added to company calendar

---

## 🚀 Tech Stack

- [Google Apps Script](https://developers.google.com/apps-script)
- Google Sheets (Backend database)
- Google Workspace (Gmail, Calendar)
- HTML + Bootstrap 5 (Email & UI templates)
- JavaScript (Client-side interaction)

---

## 📁 Project Structure

| File | Description |
|------|-------------|
| `Code.gs` | Main backend logic (approvals, notifications, data handling) |
| `form.html` | Leave request submission form |
| `dashboard.html` | Approver dashboard to review pending requests |
| `emailtemplate.html` | Approval notification email template |
| `finalNotification.html` | Final decision email template (approved/rejected) |
| `result.html` | Action confirmation page |
| `tracking.html` | Request tracking page for employees |
| `errorToken.html` | Token validation error page |
| `accessDenied.html` | Unauthorized access page |
| `cancelResult.html` | Request cancellation result page |
| `processingError.html` | Generic error page |
| `reportingEmail.html` | Reporting team notification email |
| `rejectWithNotes.html` | Rejection notes input page |
| `terms.html` | Terms of service page |
| `maintenance.html` | Maintenance mode page |

---

## 🚀 Setup Instructions

1. **Create a new Google Sheet**
2. **Rename the first sheet** to `Requests`
3. **Create a `Settings` sheet** with configuration (SPV emails, HR email, GM email, etc.)
4. **Open Apps Script editor**: `Extensions` → `Apps Script`
5. **Copy all `.gs` and `.html` files** into your Apps Script project
6. **Enable required APIs** in Apps Script:
   - People API (for profile names/photos)
   - Calendar API (for event creation)
7. **Deploy as Web App**:
   - Execute as: `User accessing the web app`
   - Access: `Anyone` (or restrict as needed)

---

## 🔒 Security Features

- **One-time approval tokens** — Each approval link is unique and expires after use
- **Email validation** — Only authorized users can view their own requests
- **Token expiration** — Used tokens are invalidated immediately
- **Secure error handling** — Graceful error pages without exposing internals

---

## 📦 Recent Updates (v4.1)

- ✅ Half-day leave display with explicit "(Half-Day)" indicator
- ✅ Requester profile photos on dashboard (with fallback avatars)
- ✅ Extracted HTML templates for easier maintenance
- ✅ Bootstrap modals for professional alerts
- ✅ Standardized message formatting across all notifications
- ✅ Mobile-responsive dashboard and forms
- ✅ Cleaned up logging for production readiness

---

## 📝 License

This project is proprietary to **ONEderland Enterprise**. Contact the maintainer for usage permissions.

---

## 👨‍💻 Maintainers

- **Iyyan Anugrah** (Creator)

For support, please contact the IT team.