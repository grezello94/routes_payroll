# Routes Payroll

Routes Payroll is a multi-company payroll app for hotel and restaurant operations.
It includes:

- admin registration and sign-in
- employee management
- monthly payroll register
- payslip preview/export/share
- leave, resume, and termination tracking
- company settings and logo management
- JSON backup/restore
- cloud-first Supabase support

The app is built for a payroll workflow where salary is paid on the 10th for the previous month.
Example: in April, the app can open March payroll by default.

## Stack

- Frontend: plain HTML, CSS, and JavaScript
- Backend: Node.js + Express
- Default cloud backend: Supabase
- Auth: app JWT session over Supabase Auth or Firebase Auth

## What The App Does

- Supports multiple companies in one workspace
- Stores employee master data separately from monthly payroll entries
- Calculates monthly gross, deductions, advance remained, and net pay
- Hides `On Leave` and `Terminated` employees from the Payroll Register
- Keeps leave, resume, and termination history in Reports
- Lets `Resumed Work` employees come back into payroll automatically
- Shows payroll by month and company
- Supports a previous-month payroll cycle by default

## Main Screens

### Employee Management

This is the employee master list.

- Shows all employees, including `Working`, `On Leave`, and `Terminated`
- Lets you edit employee details, status, leave dates, resume date, and termination date
- Shows advance remained in the employee table

### Employee Payroll Register

This is the monthly payroll working screen.

- Only active payroll employees appear here
- `On Leave` employees do not appear
- `Terminated` employees do not appear
- `Resumed Work` employees appear again in payroll
- Payroll values are stored month-wise

### Reports

Includes a `Leave / Resume Work` report.

It shows:

- employee ID
- employee name
- current status
- leave from date
- resumed on date
- terminated on date

This section is intended to keep status history visible in an organized way.

### Settings

Settings currently include:

- account details
- email verification action
- company name update
- payroll cycle mode
- designation presets
- company logo upload
- password change
- legacy payroll import

## Payroll Cycle Logic

This app supports the hotel-style payroll cycle where salary is paid on the 10th for the previous month.

Example:

- current calendar month: April
- payroll worked on in the app: March
- salary paid on: April 10

This is controlled from Settings under `Payroll Cycle`.

Available modes:

- `Previous Month (salary paid on 10th)`
- `Current Month`

By default, the app is set up to open the previous month.

## Employee Status Logic

### Working

- visible in Employee Management
- visible in Payroll Register

### On Leave

- visible in Employee Management
- hidden from Payroll Register
- leave dates remain visible in Reports

### Resumed Work

- visible in Employee Management
- visible in Payroll Register
- treated like active working mode for display
- used for payroll proration based on resume date

### Terminated

- visible in Employee Management
- hidden from Payroll Register
- termination date remains visible in Reports

## Resume-Date Payroll Logic

If an employee is marked as `Resumed Work`, the app uses the `Resumed On` date for the selected payroll month.

The payroll calculation treats the days before that resume date as absent for that month.

That means:

- payout is generated only for the days worked after resuming
- manual `Days Absent` can still be added on top if needed

## Advance Logic

The employee list shows `Advance Remained (₹)`.

This comes from the latest available payroll month for that employee.

Behavior:

- if payroll exists, the latest `advanceRemained` is shown
- if no payroll exists yet, the employee opening advance is used

In payroll:

- `Old Advance Taken` is the carried advance for that payroll month
- `Extra Advance Added` adds more advance in the month
- `Deduction Entered` reduces advance
- `Advance Remained` is what stays after deduction

## Payroll Formula

- `Gross Salary = Present Salary + Increment`
- `Total Advance = Old Advance Taken + Extra Advance Added`
- `Prorated Absence Deduction = (Days Absent / 30) × Gross Salary`
- `Deduction Applied = min(Deduction Entered, Total Advance)`
- `Advance Remained = Total Advance - Deduction Applied`
- `Net Salary = Gross Salary - Deduction Applied - Prorated Absence Deduction`

## Quick Start

```bash
cd /Users/grezello/Desktop/routes_payroll
npm install
```

Then set up Supabase and run the app.

## Supabase Setup

Supabase is the recommended backend.

### 1. Create a Supabase project

In Supabase, create a new project.

### 2. Enable Email auth

In Supabase dashboard:

- open `Authentication`
- enable `Email` sign-in

### 3. Create the database tables

Run the SQL from [supabase-setup.sql](/Users/grezello/Desktop/routes_payroll/supabase-setup.sql) in the Supabase SQL editor.

That creates these tables:

- `users`
- `companies`
- `designation_presets`
- `employees`
- `payroll_entries`
- `payroll_reports`

Important:

- if any of these tables are missing, Supabase startup can fail or some features may fall back to compatibility mode
- `payroll_reports` is used for generated payroll report snapshots
- if your older Supabase project does not have `payroll_reports`, the app will still fall back to legacy report storage, but you should still run the latest [supabase-setup.sql](/Users/grezello/Desktop/routes_payroll/supabase-setup.sql)

### 4. Create `.env.local`

Use [.env.supabase.example](/Users/grezello/Desktop/routes_payroll/.env.supabase.example) as the template.

Create `.env.local` in the project root:

```bash
DB_PROVIDER=supabase
SUPABASE_URL=https://your-project-ref.supabase.co
SUPABASE_ANON_KEY=your-anon-key
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
```

Important:

- do not commit `.env.local`
- the repo already ignores it in `.gitignore`

### 5. Start the app

Option A:

```bash
npm start
```

Option B on macOS:

```bash
./start-routes-payroll.command
```

The launcher loads `.env.local` automatically.

### 6. Open the app

```text
http://127.0.0.1:5501
```

## Deploy To Vercel

This repo is now configured so Vercel can run the Express app through a serverless entrypoint.

Files added for deployment:

- [vercel.json](/Users/grezello/Desktop/routes_payroll/vercel.json)
- [api/index.js](/Users/grezello/Desktop/routes_payroll/api/index.js)

### Recommended production backend

Use `Supabase` on Vercel.

Why:

- Vercel functions are stateless
- backup files written by the app are temporary on Vercel

The app now writes monthly emergency backups to `/tmp` on Vercel so requests do not fail, but `/tmp` is ephemeral and should not be treated as durable storage.

### Vercel environment variables

Set these in the Vercel project settings:

```text
DB_PROVIDER=supabase
SUPABASE_URL=https://your-project-ref.supabase.co
SUPABASE_ANON_KEY=your-anon-key
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
JWT_SECRET=replace-this-with-a-long-random-secret
```

Optional mail settings if you use email verification or password reset:

```text
SMTP_HOST=
SMTP_PORT=587
SMTP_USER=
SMTP_PASS=
SMTP_FROM=
SMTP_SECURE=false
```

Optional:

```text
SUPABASE_REQUEST_TIMEOUT_MS=8000
FIREBASE_WEB_API_KEY=your-firebase-web-api-key
```

### Deploy steps

1. Push this project to GitHub.
2. Import the GitHub repo into Vercel.
3. Add the environment variables above in Vercel.
4. Deploy.

If Vercel asks for build settings, the defaults are fine for this project.

### After deploy

- open the deployed URL
- register or sign in
- confirm `GET /api/health` returns a healthy response
- verify Supabase tables were created from [supabase-setup.sql](/Users/grezello/Desktop/routes_payroll/supabase-setup.sql)

## Local Run Commands

### Start

```bash
npm start
```

### Start and open on macOS

```bash
./start-routes-payroll.command
```

### Stop the local server

```bash
./stop-routes-payroll.command
```

### Syntax check

```bash
npm run check
```

## First-Time App Setup

On first launch:

1. Register the company name
2. Register admin email
3. Register username
4. Register password
5. Verify the email

After that, use Sign In normally.

## Company Behavior

- company selection happens from the top toolbar
- payroll, employees, reports, and settings are all company-scoped
- company logo is used in payslip preview and print
- company name can be edited from Settings

## Legacy Import

The app supports importing older payroll sheets from:

- `.xlsx`
- `.xls`
- `.csv`
- `.xml`

The importer:

- analyzes columns
- shows import progress from `1%` to `100%`
- shows success or exact failure reason
- infers field mappings from common header and value patterns
- imports employees, designations, and payroll rows

If import completes:

- it shows `Import successful!`

If import fails:

- it shows the reason

## Backup And Restore

The app supports JSON backup and restore.

Available from the `Actions` menu:

- `Backup DB JSON`
- `Restore DB JSON`

Supported restore payloads:

- new multi-company payload
- legacy month-only payload

## Important API Routes

### Auth

- `GET /api/auth/bootstrap`
- `POST /api/auth/register`
- `POST /api/auth/login`
- `GET /api/auth/me`
- `POST /api/auth/send-email-verification`
- `POST /api/auth/request-password-reset`
- `POST /api/auth/reset-password-with-token`
- `POST /api/auth/change-password`

### Companies

- `GET /api/companies`
- `POST /api/companies`
- `PUT /api/companies/:id`
- `PUT /api/companies/:id/logo`

### Employees

- `GET /api/employees?companyId=<id>`
- `POST /api/employees`
- `PUT /api/employees/:id`
- `DELETE /api/employees/:id?companyId=<id>`

### Payroll

- `GET /api/payroll/:month?companyId=<id>`
- `PUT /api/payroll/:month?companyId=<id>`
- `GET /api/payroll/all`
- `POST /api/payroll/restore`

### Settings

- `GET /api/settings/designations?companyId=<id>`
- `POST /api/settings/designations`
- `DELETE /api/settings/designations/:id?companyId=<id>`

## Files To Know

- [server.js](/Users/grezello/Desktop/routes_payroll/server.js): Express backend and APIs
- [datastore.js](/Users/grezello/Desktop/routes_payroll/datastore.js): storage providers for Supabase, Firebase, and SQLite
- [app.js](/Users/grezello/Desktop/routes_payroll/app.js): frontend app logic
- [index.html](/Users/grezello/Desktop/routes_payroll/index.html): app structure
- [styles.css](/Users/grezello/Desktop/routes_payroll/styles.css): UI styling
- [supabase-setup.sql](/Users/grezello/Desktop/routes_payroll/supabase-setup.sql): Supabase schema setup
- [.env.supabase.example](/Users/grezello/Desktop/routes_payroll/.env.supabase.example): Supabase env template
- [start-routes-payroll.command](/Users/grezello/Desktop/routes_payroll/start-routes-payroll.command): local launcher
- [stop-routes-payroll.command](/Users/grezello/Desktop/routes_payroll/stop-routes-payroll.command): local stop script

## Firebase And SQLite Notes

Supabase is the preferred backend now.

Firebase support still exists as a fallback for older setups.

## Security Notes

- `.env.local` should stay local
- do not commit service-role keys
- do not commit Firebase service account files
- set a custom `JWT_SECRET` in production

## Troubleshooting

### App says old backend is running

Restart with:

```bash
cd /Users/grezello/Desktop/routes_payroll
./stop-routes-payroll.command
./start-routes-payroll.command
```

Then hard refresh the browser:

```text
Cmd + Shift + R
```

### Supabase is not connecting

Check:

- `SUPABASE_URL`
- `SUPABASE_ANON_KEY`
- `SUPABASE_SERVICE_ROLE_KEY`
- whether `.env.local` exists
- whether the project URL matches the real Supabase API settings

### Payroll month is wrong

Open `Settings` and change `Payroll Cycle`.

### Leave or terminated employee is still visible in payroll

Restart the backend and refresh the browser so the latest filtering logic is used.

## Current Recommended Workflow

1. Keep employee master data updated in Employee Management.
2. Use status fields for leave, resume, and termination dates.
3. Review status history in Reports.
4. Open Payroll Register for the target payroll month.
5. Complete payroll values and review computed payout.
6. Export payslips, Excel, CSV, or backup JSON as needed.
