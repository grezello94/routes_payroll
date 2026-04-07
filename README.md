# Routes Payroll

Professional multi-company payroll PWA with backend auth, Supabase cloud storage, and export tools.

## Stack

- Frontend: HTML/CSS/JS (PWA-capable)
- Backend: Node.js + Express
- Database: Supabase by default when configured, otherwise Firebase or SQLite fallback
- Auth: Supabase Auth or Firebase Auth behind the same JWT-based app login

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

On macOS you can also double-click:

```bash
start-routes-payroll.command
```

That starts the backend if needed and opens the app automatically.
If `.env.local` exists in the project root, the launcher will load it before starting.

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

## Supabase (Recommended)

Supabase is now the recommended cloud backend for this project.

Set environment variables and run:

```bash
DB_PROVIDER=supabase \
SUPABASE_URL="https://YOUR_PROJECT.supabase.co" \
SUPABASE_ANON_KEY="your-anon-key" \
SUPABASE_SERVICE_ROLE_KEY="your-service-role-key" \
npm start
```

Notes:

- Run the SQL in [supabase-setup.sql](/Users/grezello/Desktop/routes_payroll/supabase-setup.sql) inside the Supabase SQL editor before first start.
- Copy [.env.supabase.example](/Users/grezello/Desktop/routes_payroll/.env.supabase.example) to `.env.local` and fill in your real keys for easier local startup.
- `DB_PROVIDER=supabase` explicitly forces Supabase mode.
- When `SUPABASE_URL` and `SUPABASE_SERVICE_ROLE_KEY` are present, the datastore will prefer Supabase automatically.
- `SUPABASE_ANON_KEY` is used for password sign-in requests.
- `DB_PROVIDER=sqlite` is only for an intentional local-only fallback.
- Email/password users live in Supabase Auth.
- Username, email verification state, companies, employees, designations, and payroll rows live in Supabase tables.

## Firebase Database (Legacy / Fallback)

Firebase support is still present as a fallback while migrating.

```bash
DB_PROVIDER=firebase \
FIREBASE_PROJECT_ID="your-project-id" \
FIREBASE_SERVICE_ACCOUNT_JSON='{"type":"service_account", ...}' \
npm start
```

This repo can connect to Firebase using either:
- `FIREBASE_SERVICE_ACCOUNT_JSON`
- `GOOGLE_APPLICATION_CREDENTIALS`
- A local service account file at the project root such as `firebase-service-account.json` or `routespayroll-firebase-adminsdk-*.json`

To sync the current local admin account into Firestore so login works the same across systems:

```bash
npm run sync:firebase-auth
```

## Optional Security Hardening

Set custom JWT secret before start:

```bash
JWT_SECRET="replace-with-a-strong-random-secret" npm start
```
