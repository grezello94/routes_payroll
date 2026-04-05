# Routes Payroll

Professional multi-company payroll PWA with backend auth, SQLite storage, branded payslips, and export tools.

## Stack

- Frontend: HTML/CSS/JS (PWA-capable)
- Backend: Node.js + Express
- Database: SQLite (`data/payroll.db`) or Firebase Firestore
- Auth: JWT-based admin login

## Key Features

- Secure admin setup/login
- Multi-company payroll workspace
  - Create company with `name` + `logo`
  - Switch active company from the top toolbar
  - Data is isolated per company + month
- Month-wise payroll register with auto-save
- Payslip preview with company branding
  - Company name + logo rendered in payslip
  - Print/Save PDF
  - TXT download
  - Share: WhatsApp, Messenger, Web Share, copy text
- Exports
  - CSV (monthly)
  - Excel `.xls` (monthly, highlighted professional format)
- Backup/restore (JSON)
  - Supports new multi-company payload format
  - Backward-compatible with legacy month-only payloads
- Installable app (service worker + manifest)

## Run

```bash
npm install
npm start
```

Open: `http://127.0.0.1:5501`

## Quick Start (Do This)

```bash
cd /Users/grezello/Desktop/routes_payroll
npm install
npm start
```

Then open:

- Main app: `http://127.0.0.1:5501`
- HRM mockup interface: `http://127.0.0.1:5501/hrm_interface_mockup.html`

## How Multi-Company Works

1. Use `New Company` to create a company with logo.
2. Select company from the `Company` dropdown.
3. All month reads/writes are scoped by selected company.
4. Payslip and exports reflect the selected company branding.

## API Endpoints

- `GET /api/health`
- `GET /api/auth/bootstrap`
- `POST /api/auth/register`
- `POST /api/auth/login`
- `GET /api/auth/me`

### Company APIs

- `GET /api/companies`
- `POST /api/companies`
  - Body:
    - `name: string`
    - `logoDataUrl: string` (optional, image data URL)

### Payroll APIs

- `GET /api/payroll/:month?companyId=<id>`
- `PUT /api/payroll/:month?companyId=<id>`
  - Body: `{ records: [...] }`
- `GET /api/payroll/all`
  - Returns multi-company backup shape:
    - `{ companies: [...], entries: [...] }`
- `POST /api/payroll/restore`
  - Accepts:
    - New format: `{ companies: [...], entries: [...] }`
    - Legacy format: `{ months: { "YYYY-MM": [...] } }`

## Payroll Formulas

- `Gross Salary = Present Salary + Increment`
- `Total Advance = Old Advance Taken + Extra Advance Added`
- `Prorated Absence Deduction = (Days Absent / 30) × Gross Salary`
- `Deduction Applied = MIN(Deduction Entered, Total Advance)`
- `Advance Remained = Total Advance − Deduction Applied`
- `Net Salary = Gross Salary − Deduction Applied − Prorated Absence Deduction`

## Data Notes

- SQLite DB file: `data/payroll.db`
- JWT token is stored in browser `localStorage`
- Company logos are stored as image data URLs in SQLite

## Firebase Database (Optional)

You can switch backend storage from SQLite to Firebase Firestore.

Set environment variables and run:

```bash
DB_PROVIDER=firebase \
FIREBASE_PROJECT_ID="your-project-id" \
FIREBASE_SERVICE_ACCOUNT_JSON='{"type":"service_account", ...}' \
npm start
```

Notes:

- `DB_PROVIDER=firebase` enables Firebase mode.
- `FIREBASE_SERVICE_ACCOUNT_JSON` should be a full service account JSON string.
- You can use `GOOGLE_APPLICATION_CREDENTIALS` instead of `FIREBASE_SERVICE_ACCOUNT_JSON` if preferred.

## Optional Security Hardening

Set custom JWT secret before start:

```bash
JWT_SECRET="replace-with-a-strong-random-secret" npm start
```
