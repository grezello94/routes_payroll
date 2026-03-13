# Routes Payroll (Full Backend + Database)

Routes Payroll is now upgraded to a full-stack PWA:
- Frontend: HTML/CSS/JS
- Backend API: Node.js + Express
- Database: SQLite (`data/payroll.db`)
- Auth: JWT-based admin login
- PWA: installable with offline static asset caching

## Features

- Secure admin setup/login (server-side credentials)
- Year-month payroll management stored in SQLite
- Full payroll fields + formulas
- Payslip preview + print/save PDF + TXT export
- Share via WhatsApp / Messenger / Web Share
- CSV export
- Full DB backup/restore as JSON
- White/black professional UI with color-coded payroll fields

## Run

From `C:\Users\Grezello\Desktop\PAYROLL`:

```powershell
npm install
npm start
```

Open:

`http://127.0.0.1:5501`

## API Endpoints

- `GET /api/health`
- `GET /api/auth/bootstrap`
- `POST /api/auth/register`
- `POST /api/auth/login`
- `GET /api/auth/me`
- `GET /api/payroll/:month`
- `PUT /api/payroll/:month`
- `GET /api/payroll/all`
- `POST /api/payroll/restore`

## Payroll Formulas

- `Gross Salary = Present Salary + Increment`
- `Total Advance = Old Advance Taken + Extra Advance Added`
- `Prorated Absence Deduction = (Days Absent / 30) × Gross Salary`
- `Deduction Applied = MIN(Deduction Entered, Total Advance)`
- `Advance Remained = Total Advance − Deduction Applied`
- `Net Salary = Gross Salary − Deduction Applied − Prorated Absence Deduction`

## Data

- SQLite file path: `data/payroll.db`
- Auth token: browser `localStorage`

## Optional Security Hardening

Set a custom JWT secret before starting:

```powershell
$env:JWT_SECRET = "replace-with-a-strong-random-secret"
npm start
```
