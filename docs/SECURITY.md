# Security and privacy checklist

## Current status

This repository is public. Treat all source code and committed content as public. Never commit real financial data, passwords, OAuth credentials, API tokens, session secrets, database exports, or .dev.vars files.

The dashboard is a visual prototype with hard-coded sample data. The Worker has a health endpoint and Better Auth route wiring, but authentication and database behavior have **not yet been built and verified end-to-end**. Do not use this version for real financial records or advertise it as production-ready.

## Required before launch

- [ ] Confirm the Better Auth schema against the exact installed package version.
- [ ] Configure a unique BETTER_AUTH_SECRET using Cloudflare secrets.
- [ ] Configure the exact BETTER_AUTH_URL and trusted origin.
- [ ] Test account creation, login, logout, session expiration, and rate limiting.
- [ ] Add email verification and password-reset delivery, or choose and configure a supported external sign-in provider.
- [ ] Require a validated session for every private finance API route.
- [ ] Derive user_id from the verified session on the server; never trust a client-supplied user ID.
- [ ] Scope every SELECT, INSERT, UPDATE, and DELETE to the authenticated user. Check ownership for referenced account/category/goal IDs too.
- [ ] Validate request bodies, amounts, dates, and pagination on the server.
- [ ] Add CSRF/origin protections appropriate to the final cookie and API configuration.
- [ ] Add automated tests proving one account cannot read, edit, export, or delete another account's data.
- [ ] Ensure reports are generated from the signed-in user's filtered records only.
- [ ] Verify backups, retention, deletion/export requests, and a recovery procedure.
- [ ] Confirm deployment configuration and logs never expose secrets or sensitive financial values.

## Data design rules

- Store money as integer centavos, not floating-point values.
- Keep financial data in D1, never in GitHub.
- Treat transfers between a user's own accounts as transfers, not income or spending.
- Use server-side authorization for all financial actions. Hiding a button in the interface is not access control.
- A public repository is acceptable for public source code only; it is not a security boundary for user data.