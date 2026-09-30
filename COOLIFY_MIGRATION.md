# Vercel To Coolify On Contabo

This project is ready to run as a Docker app in Coolify.

## What Changes From Vercel

- Vercel ran the app through `api/index.js` as a serverless Express function.
- Coolify should run the app as a normal long-running Node server with `npm start`.
- Vercel rewrites are not needed because Express already serves static files and API routes.
- Vercel Cron must be replaced with a Coolify scheduled job or an external uptime monitor.

## Recommended Production Setup

Keep Supabase as the production database while moving the app hosting to Contabo/Coolify.

This avoids moving payroll data into a local container database and makes migration safer:

- app server: Coolify on Contabo
- database: existing Supabase project
- env vars: copied from Vercel to Coolify
- domain: moved from Vercel DNS target to Coolify proxy

## Files Added For Coolify

- `Dockerfile`
- `.dockerignore`
- `.env.coolify.example`

## Coolify Deployment Steps

1. Push this repo to GitHub or any Git provider Coolify can access.
2. In Coolify, create a new resource from the repository.
3. Select Dockerfile build.
4. Set the exposed port to `5501`.
5. Add environment variables from `.env.coolify.example`.
6. Set the domain in Coolify.
7. Deploy.
8. Open `/api/health` on the deployed domain.

Expected healthy response:

```json
{
  "ok": true,
  "pending": false,
  "degraded": false,
  "provider": "supabase"
}
```

## Environment Variables To Copy From Vercel

Required:

```text
DB_PROVIDER=supabase
SUPABASE_URL=...
SUPABASE_ANON_KEY=...
SUPABASE_SERVICE_ROLE_KEY=...
JWT_SECRET=...
CRON_SECRET=...
```

Optional:

```text
SMTP_HOST=...
SMTP_PORT=587
SMTP_USER=...
SMTP_PASS=...
SMTP_FROM=...
SMTP_SECURE=false
SUPABASE_REQUEST_TIMEOUT_MS=30000
FIREBASE_WEB_API_KEY=...
FIREBASE_SERVICE_ACCOUNT_JSON=...
FIREBASE_PROJECT_ID=...
```

## Persistent Storage

If `DB_PROVIDER=supabase`, payroll data is stored in Supabase. A volume is not required for normal app data.

If you choose SQLite instead, add a persistent volume in Coolify:

```text
/app/data
```

Without that volume, SQLite data and monthly backup files can be lost when the container is rebuilt.

## Replacing Vercel Cron

Vercel called:

```text
GET /api/keepalive
Authorization: Bearer <CRON_SECRET>
```

In Coolify, create a scheduled task or use an external uptime monitor that sends the same request once per day.

Example command:

```bash
curl -fsS -H "Authorization: Bearer $CRON_SECRET" "https://your-domain.com/api/keepalive"
```

## DNS Cutover

1. Deploy successfully on a temporary Coolify URL first.
2. Confirm `/api/health` is healthy.
3. Confirm login, company list, employee list, and payroll month load.
4. In your DNS provider, point the app domain to the Contabo server as Coolify instructs.
5. Enable SSL in Coolify.
6. Keep the Vercel deployment untouched until the Coolify domain is verified.

## Rollback

If anything fails after DNS cutover, point DNS back to Vercel. Because the recommended setup keeps Supabase unchanged, rollback does not require restoring app data.

## Final Verification Checklist

- `/api/health` returns `provider: "supabase"` and `degraded: false`
- registration or login works
- current company loads
- employee management loads
- payroll register loads for the target month
- payslip preview/export works
- password reset or email verification works if SMTP is configured
- scheduled keepalive returns `ok: true`
