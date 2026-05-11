# Google Service Account Setup — Step by Step

**Goal:** Give the Streamlit app a "robot Google account" (called a service account) that can read and write the canopy quote tracker Google Sheet on Ali's Drive.

**Time required:** ~15 minutes, one-time setup.

**Sheet you're connecting:** `https://docs.google.com/spreadsheets/d/1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w`

---

## What you'll end up with

By the end of this guide you'll have:

1. A Google Cloud project (free tier is fine — no billing required)
2. A service account inside that project with an email like `canopy-tracker@your-project-id.iam.gserviceaccount.com`
3. A JSON key file downloaded to your computer (the credentials)
4. Sheets and Drive APIs enabled on the project
5. The Google Sheet shared with the service account's email (Editor access)
6. The JSON key pasted into Streamlit Cloud's secrets manager

Then the app can write to the sheet automatically every time a proposal is generated.

---

## Step 1 — Create a Google Cloud project

1. Open https://console.cloud.google.com in the same browser where you're signed into the Google account that owns the canopy tracker sheet.
2. At the top of the page, click the project dropdown (it says "Select a project" or shows an existing project name).
3. Click **NEW PROJECT** (top right of the dropdown panel).
4. Project name: `Canopy Tracker` (or anything you like).
5. Leave Organization as "No organization" if asked.
6. Click **CREATE**. Wait ~10 seconds for the project to provision.
7. Make sure the project dropdown at the top now shows `Canopy Tracker`. If not, click the dropdown and select it.

---

## Step 2 — Enable the Sheets and Drive APIs

The service account needs explicit permission to use Google Sheets and Google Drive APIs.

1. In the search bar at the top of Google Cloud Console, type `Google Sheets API` and click the matching result.
2. Click the blue **ENABLE** button. Wait a few seconds.
3. Use the search bar again to find `Google Drive API` and click the matching result.
4. Click **ENABLE** again.

(Drive API is needed because gspread uses it to open sheets by ID.)

---

## Step 3 — Create the service account

1. In the search bar, type `Service Accounts` and click the result under "IAM & Admin".
2. Click **+ CREATE SERVICE ACCOUNT** at the top of the page.
3. Service account name: `canopy-tracker`
4. Service account ID: will auto-fill as `canopy-tracker` — leave it.
5. Description: `Writes to canopy quote tracker sheet from Streamlit app` (optional).
6. Click **CREATE AND CONTINUE**.
7. **Grant this service account access to project** screen — leave the role field blank and click **CONTINUE**. (The service account doesn't need any project-level roles; it only needs sheet-level access, which we grant in Step 5.)
8. **Grant users access to this service account** screen — leave both fields blank and click **DONE**.

You'll be returned to the Service Accounts list. You should see `canopy-tracker@<your-project-id>.iam.gserviceaccount.com` in the list. **Copy that email address** — you'll need it in Step 5.

---

## Step 4 — Generate and download the JSON key

1. Click on the `canopy-tracker` service account row to open its detail page.
2. Click the **KEYS** tab at the top.
3. Click **ADD KEY → Create new key**.
4. Choose **JSON** as the key type.
5. Click **CREATE**.

A JSON file will download to your computer automatically (something like `canopy-tracker-abc123.json`). **This file is the password.** Treat it like a password:

- Don't email it
- Don't post it in Slack
- Don't commit it to GitHub
- Save it somewhere private (e.g., a password manager, or your local Downloads folder for now)

If you ever lose it, no problem — come back to this Keys tab and create a new one. The old one can be deleted any time.

---

## Step 5 — Share the Google Sheet with the service account

This is the easy-to-forget step. The service account exists but has no access to your sheet yet.

1. Open the canopy tracker sheet: https://docs.google.com/spreadsheets/d/1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w
2. Click the green **Share** button (top right).
3. In the "Add people, groups, and calendar events" field, paste the service account email from Step 3 (`canopy-tracker@<your-project-id>.iam.gserviceaccount.com`).
4. Make sure the role dropdown next to it is set to **Editor** (not Viewer — the app needs to write rows).
5. **Uncheck "Notify people"** — there's no inbox to notify.
6. Click **Share**.

The service account now has Editor access to this sheet only.

---

## Step 6 — Give the JSON to your web developer for Streamlit Cloud

The web developer needs to paste the JSON contents into Streamlit Cloud's secrets manager so the app can read them.

**Forward the downloaded JSON file to your web developer with this note:**

> Hi — here's the service account JSON for the canopy quote tracker. Please paste it into the Streamlit Cloud app's Settings → Secrets, in TOML format, under a section named exactly `[gcp_service_account]`. The fields in the JSON map one-to-one to TOML keys. The `private_key` field needs to keep its `\n` newline characters intact (TOML triple-quoted string or escaped string both work). After saving, Streamlit will redeploy automatically. Confirmation: the admin panel in the app should now show "Source: Google Sheets" instead of "Source: Local CSV".

Your web developer has done this kind of paste before — there's a worked example in `DEPLOY_NOTES.md` showing the exact TOML format.

---

## Step 7 — Verify it works

After your web developer has pasted the secret and the Streamlit app has redeployed:

1. Open https://www.alihusain.me
2. Fill in any test scenario and click **Calculate Canopy Price**.
3. Click **📄 Generate Word Proposal**. You should see a success message with a quote number.
4. **Open the canopy tracker sheet directly** (https://docs.google.com/spreadsheets/d/1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w) — you should see a new row at the bottom with the quote details.
5. Scroll to the bottom of the configurator page and expand the **🔒 Profitability Tracker (admin)** section. Enter the admin password. The header at the top should say **Source: Google Sheets** with a link.

If the admin panel says "Source: Local CSV (fallback — Google Sheets not configured)", the secret didn't take — ask your web developer to double-check the TOML section is named exactly `[gcp_service_account]` (no spaces, no typos, square brackets included).

---

## Troubleshooting

- **"PERMISSION_DENIED" warning in the app:** the sheet wasn't shared with the service account email. Repeat Step 5.
- **"API has not been used... or it is disabled":** Sheets or Drive API wasn't enabled on the project. Repeat Step 2.
- **App still uses local CSV after deploy:** Streamlit secrets aren't being read. Confirm the TOML section name is exactly `[gcp_service_account]` and that the deploy actually picked up the new secret (sometimes you have to "Reboot" the app from the Streamlit Cloud dashboard).
- **Lost the JSON key:** create a new one in Step 4. Have web dev paste the new one. Delete the old key from the Keys tab to revoke it.

---

## Cost

Free. Service accounts and the Google Sheets / Drive APIs are free for normal use. The canopy tracker writes one row per proposal — well under any rate limit. No billing account is required for this Cloud project.
