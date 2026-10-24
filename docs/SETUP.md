# Budgetwise setup

This branch contains the initial UI and Cloudflare Worker scaffold. The app is **not production-ready**: the dashboard uses sample data and must not be treated as a real financial record system.

## Stack

- React + TypeScript + Vite
- Cloudflare Workers and Static Assets
- Cloudflare D1 (SQLite)
- Better Auth for email/password sessions
- ExcelJS for xlsx report generation

## Local development

1. Install Node.js LTS and npm.
2. Clone the repository and check out feature/personal-budget-tracker.
3. Run npm install.
4. Run npm run dev to open the UI.
5. Run npm run typecheck and npm run build before committing.

Do not place real transactions or account information in sample data, source files, issue descriptions, or commits.

## Cloudflare setup (do not deploy until configuration is complete)

1. Create a **new D1 database** named budgetwise-db in the intended Cloudflare account.
2. Replace REPLACE_WITH_CLOUDFLARE_D1_DATABASE_ID in wrangler.jsonc with the actual D1 database ID.
3. Generate a unique random authentication secret and add it using Cloudflare Worker secrets as BETTER_AUTH_SECRET. Never commit it.
4. Add BETTER_AUTH_URL as a Worker variable once the workers.dev address is known. Keep it equal to the exact deployed origin.
5. Review the Better Auth schema for the installed version, then apply the D1 migrations to the **intended database only**.
6. Run the app locally and verify sign-up, sign-in, sign-out, session expiry, and that a second user cannot access the first user's records.
7. Only after those checks pass, deploy to the temporary workers.dev address. A custom domain can be attached later.

Suggested commands after installing dependencies and replacing the database ID:
- npx wrangler d1 migrations list budgetwise-db --remote
- npx wrangler d1 migrations apply budgetwise-db --remote
- npm run build
- npx wrangler deploy

Review each command and its target before running it. Do not apply migrations to a production database until the migration contents and target database have been checked.

## Reporting

The prototype's Excel export currently exports **sample values only**. The real export must be connected to authenticated, user-scoped database queries before this feature is used with personal data.

## Deployment status

No D1 database has been created or connected, no secret has been set, and no Cloudflare deployment has been performed from this branch.