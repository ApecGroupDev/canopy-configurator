# Canopy Configurator — Deploy Notes (v8 update)

**App:** Streamlit Python web app for canopy pricing + proposal generation.
**Production target:** `www.alihusain.me`
**Update type:** Existing app replacement. Adds: Word proposal download button, profitability tracker (CSV), admin panel.

---

## Files in this bundle (6)

All six files must live **in the same directory** on the server:

| File | Purpose |
|------|---------|
| `canopy_configurator.py` | Main Streamlit app (entry point) — v8 |
| `proposal_writer.py` | Word proposal generator (imported by configurator) — **NEW MODULE NAME** |
| `canopy_config.xlsx` | All pricing rates, cheat sheets, defaults — read at startup |
| `Apec Imaging Logo.jpg` | Header logo when APEC is selected |
| `GEO Canopies logo.jpg` | Header logo when GEO is selected |
| `requirements.txt` | Python dependencies |

---

## Important: remove the old `proposal_builder.py`

The previous bundle had a module named `proposal_builder.py`. The new bundle replaces it with **`proposal_writer.py`** (different filename, same role). When deploying:

1. **Delete `proposal_builder.py`** from the existing app directory on the server.
2. **Add `proposal_writer.py`** from this bundle.

The configurator's import statement is wired to `proposal_writer`, not `proposal_builder`. Leaving the old file in place won't break anything but is dead weight.

---

## Requirements

- Python 3.9 or newer
- `pip` for installing dependencies

`requirements.txt` includes: `streamlit`, `openpyxl`, `python-docx`.

---

## Install

From the app directory:

```bash
pip install -r requirements.txt
```

---

## Run

```bash
streamlit run canopy_configurator.py
```

Default port: `8501`. Reverse-proxy / hosting config from the previous deploy should not need changes.

---

## What's new in this version (v8)

Three new features beyond the existing pricing calculator:

1. **Word proposal generation** — at the bottom of the results section, a "📄 Generate Word Proposal" button assigns a quote number, generates a customer-facing 2-page `.docx`, and offers it via a download button.
2. **Profitability tracker** — every proposal generation appends a row to `quote_tracker.csv` (auto-created in the app directory). Columns: Date, Quote No, Customer Name, City, Sales Rep, Grand Total, Profitability.
3. **Admin panel** — password-gated expander at the bottom of the page. Shows the tracker table + offers a CSV download. The default admin password is `profit_tracker` and is defined on line 29 of `canopy_configurator.py` (constant `ADMIN_PASSWORD`). The site owner will likely change this to something else.

---

## Persistence considerations — IMPORTANT

The `quote_tracker.csv` file is written **at runtime** in the app directory. It does not exist in the bundle; it auto-creates on first proposal generation.

For this file to be reliable across server restarts and redeploys:

- The app directory must be on **persistent storage** (not an ephemeral /tmp or container-rebuild path). If you're using a container deployment, mount a persistent volume at the app directory or write the CSV to a mounted location.
- The file should be included in the server's **nightly backup** — it contains every quote ever generated and is the only record of profitability data.
- The file **must not be web-publicly accessible** — it contains internal margin data. If you're using a reverse proxy, ensure paths like `/quote_tracker.csv` return 404 (or are otherwise blocked). Streamlit itself does not expose arbitrary files in its app directory, but a misconfigured static-file serving rule could.

If you'd rather store the CSV in a different absolute path (e.g. `/var/lib/canopy/quote_tracker.csv`), edit the `TRACKER_PATH` constant on line 33 of `canopy_configurator.py`.

---

## Replacing the previous version

1. Stop the running Streamlit process.
2. In the existing app directory: **delete** `proposal_builder.py` (the old module name).
3. **Replace** the existing `canopy_configurator.py` and `canopy_config.xlsx` and both `.jpg` logos and `requirements.txt` with the new versions from this bundle.
4. **Add** `proposal_writer.py` from this bundle.
5. Re-run: `pip install -r requirements.txt` (no new deps, but safe to refresh).
6. Restart: `streamlit run canopy_configurator.py`.
7. Confirm `quote_tracker.csv` does NOT exist yet (it will be created on first proposal). Confirm the app directory is on persistent storage.

---

## Quick verification after deploy

1. Open the live site.
2. Fill in any test scenario, click "Calculate Canopy Price" — quote results should render as before.
3. Click "📄 Generate Word Proposal" — success message appears with a quote number like `Q-20260511-001`.
4. Click "⬇️ Download Proposal (.docx)" — a Word file downloads. Open in Word: should be 2 pages, no errors.
5. SSH to the server (or check via the admin panel) — confirm `quote_tracker.csv` was created with one row.
6. Scroll to the bottom of the page, expand "🔒 Profitability Tracker (admin)", enter password `profit_tracker` — should display the row that was just logged with a CSV download button.

If any of the above fails, return the error output and we'll fix.

---

## Filenames are case- and space-sensitive

Don't rename:
- `Apec Imaging Logo.jpg`
- `GEO Canopies logo.jpg`
- `canopy_config.xlsx`
- `proposal_writer.py`

The app looks for these exact names.
