# Canopy Configurator — Deploy Notes (v9 update)

**App:** Streamlit Python web app for canopy pricing + proposal generation.
**Production target:** `www.alihusain.me` (Streamlit Community Cloud)
**Update type:** Existing app replacement.
**v9 change:** Quote tracker moved from local CSV to Google Sheets (Streamlit Cloud's filesystem is ephemeral and was wiping the CSV on every restart).

---

## Files in this bundle (6)

All six files must live **in the same directory** on the server:

| File | Purpose |
|------|---------|
| `canopy_configurator.py` | Main Streamlit app (entry point) — v9 |
| `proposal_writer.py` | Word proposal generator (imported by configurator) |
| `canopy_config.xlsx` | All pricing rates, cheat sheets, defaults — read at startup |
| `Apec Imaging Logo.jpg` | Header logo when APEC is selected |
| `GEO Canopies logo.jpg` | Header logo when GEO is selected |
| `requirements.txt` | Python dependencies (now includes gspread + google-auth) |

---

## What changed in v9

The quote tracker no longer writes to a local CSV file. Instead it reads/writes to a Google Sheet:

- **Sheet:** `https://docs.google.com/spreadsheets/d/1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w`
- **Worksheet name:** `Tracker` (auto-created on first write if missing)
- **Auth:** Google Cloud service account; credentials JSON pasted into Streamlit Cloud's secrets manager under the section `[gcp_service_account]`

The app falls back to the local CSV (`quote_tracker.csv`) if Streamlit secrets aren't configured — useful for local development. In production on Streamlit Cloud, secrets must be set or quotes will only be logged in-memory and lost on restart.

---

## Streamlit Cloud secrets — REQUIRED for production

In the Streamlit Cloud app dashboard:

1. Open the app → **Settings → Secrets**
2. Paste the entire content of the service-account JSON key file as TOML, like this:

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "abc123..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "canopy-tracker@your-project-id.iam.gserviceaccount.com"
client_id = "1234567890..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
universe_domain = "googleapis.com"
```

3. Save. Streamlit will redeploy automatically.

**The service-account creation walkthrough is in `GOOGLE_SHEETS_SETUP.md`** in this bundle.

After the secrets are saved, you also need to **share the Google Sheet** with the service account's `client_email` (the one ending in `iam.gserviceaccount.com`). Give it Editor access. Without this, the app can authenticate but can't open the sheet.

---

## Replacing the previous version

1. In the Streamlit Cloud app: **Settings → Secrets** — paste the TOML block above (after Ali completes the service-account setup).
2. In your GitHub repo backing the app, replace these files with the v9 versions:
   - `canopy_configurator.py`
   - `requirements.txt`
   - `proposal_writer.py` (unchanged from v8 but include for completeness)
   - `canopy_config.xlsx` (unchanged from v8)
   - Both `.jpg` logos (unchanged)
3. The old `quote_tracker.csv` in the repo (if any) can be deleted — it's no longer used in production.
4. Push to the branch Streamlit Cloud is watching. Streamlit will auto-redeploy.
5. Confirm the new deploy installed `gspread` and `google-auth` (check the build log).

---

## Admin password

Default admin password (line 41 of `canopy_configurator.py`): `profit_tracker`. Change this to something only Ali knows before going live.

---

## Quick verification after deploy

1. Open the live site.
2. Fill in any test scenario, click **Calculate Canopy Price**.
3. Click **📄 Generate Word Proposal** — success message with a quote number like `Q-20260511-001`.
4. Open the Google Sheet directly (`https://docs.google.com/spreadsheets/d/1pnfCv70Y4UBWw8A2Yi49QBvr2ErwRDy7raxPdkGkC1w`) — confirm a new row appeared in the `Tracker` worksheet.
5. Click **⬇️ Download Proposal (.docx)** — Word file downloads, opens cleanly in Word.
6. Scroll to the bottom of the page, expand **🔒 Profitability Tracker (admin)**, enter the admin password — should show the row that was just logged with a "Source: Google Sheets" label and a link to open the sheet.

If the admin panel says "Source: Local CSV (fallback — Google Sheets not configured)", it means the Streamlit secrets aren't being picked up. Re-check the TOML formatting and that the section is named exactly `[gcp_service_account]`.

---

## Persistence + security notes

- Google Sheets handles persistence automatically — no backups needed beyond Google's own.
- **Sheet sharing:** the sheet is shared only with Ali's personal Google account + the service-account robot email. Do not enable "Anyone with the link" sharing — it would expose every customer name and quote total publicly.
- If the service-account JSON is ever leaked, immediately rotate it: in Google Cloud Console → IAM → Service Accounts → Keys, delete the old key and create a new one. Update the Streamlit secret with the new JSON.

---

## Filenames are case- and space-sensitive

Don't rename:
- `Apec Imaging Logo.jpg`
- `GEO Canopies logo.jpg`
- `canopy_config.xlsx`
- `proposal_writer.py`

The app looks for these exact names.
