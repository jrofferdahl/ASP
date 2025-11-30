# ASP

This repository contains the Google Apps Script project exported from the Apps Script editor via clasp.

Contents
- The Apps Script source files are at the repository root (code.js, html files, appsscript.json, .clasp.json).
- Use clasp to sync with Google Apps Script:
  - Pull changes from Apps Script: `clasp pull`
  - Push local changes to Apps Script: `clasp push`
  - Open the project in the Apps Script editor: `clasp open`

CI / Deploy
- This repo includes a GitHub Actions workflow (.github/workflows/deploy.yml) that, on push to main, will:
  1) install clasp
  2) authenticate using a secret (CLASPRC_JSON)
  3) push code to Apps Script
  4) create a new version and deploy it

Secrets required for the workflow
- CLASPRC_JSON: the full contents of your local ~/.clasprc.json (created by `clasp login`).
  - Store this value as a repository Secret (Repository > Settings > Secrets and variables > Actions > New repository secret).
  - See the README section below for how to extract and create that secret.

Security note
- CLASPRC_JSON contains an OAuth refresh token for the account used to run `clasp login`. Treat it like a password. If you prefer not to use your personal tokens in CI, see the "Service account" section below for a safer alternative.

How to add the CLASPRC_JSON secret (Windows)
1. On your local machine open PowerShell and run:
   Get-Content $env:USERPROFILE\.clasprc.json -Raw
2. Copy the entire JSON output.
3. In GitHub go to your repository → Settings → Secrets and variables → Actions → New repository secret.
   - Name: CLASPRC_JSON
   - Value: (paste the full JSON you copied)
4. Save the secret.

How to add the CLASPRC_JSON secret (macOS / Linux)
1. Run:
   cat ~/.clasprc.json
2. Copy the full JSON and add it as the CLASPRC_JSON repository secret.

Alternative: Service account (recommended for orgs)
- Instead of storing your personal ~/.clasprc.json, create a Google Cloud service account with appropriate Apps Script access, generate a JSON key, and follow the clasp service account flow. If you want this option, I’ll provide exact steps for:
  - creating the service account,
  - granting it script access,
  - adding the key as a GitHub secret,
  - and adjusting the workflow to use it.
