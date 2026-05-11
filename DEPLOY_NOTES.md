# Canopy Configurator — Deploy Notes

**App:** Streamlit web app for canopy pricing.
**Current production target:** `www.alihusain.me`
**This bundle:** v6 (adds GEO Canopies logo swap based on user selection).

---

## Files in this bundle (5)

All five files must live **in the same directory** on the server:

| File | Purpose |
|------|---------|
| `canopy_configurator.py` | Main Streamlit app (entry point) |
| `canopy_config.xlsx` | All pricing rates, cheat sheets, defaults — read at startup |
| `Apec Imaging Logo.jpg` | Header logo when APEC Canopies is selected |
| `GEO Canopies logo.jpg` | Header logo when GEO Canopies is selected (**NEW in this version**) |
| `requirements.txt` | Python dependencies |

---

## Requirements

- Python 3.9 or newer
- `pip` for installing dependencies

---

## Install

From the app directory:

```bash
pip install -r requirements.txt
```

This installs Streamlit, openpyxl, and python-docx.

---

## Run

```bash
streamlit run canopy_configurator.py
```

Default port: `8501`. Reverse-proxy / hosting config from the previous deploy should not need changes — only the file bundle is updated.

---

## What's new in this version (v6)

The app now displays the correct company logo at the top based on the radio-button selection ("APEC Canopies" vs "GEO Canopies"). The only new file required on the server is `GEO Canopies logo.jpg`. Everything else is a routine update of the existing files.

---

## Replacing the previous version

1. Stop the running Streamlit process.
2. Replace all 5 files in the existing app directory with the files in this bundle.
3. Re-run: `pip install -r requirements.txt` (in case dependency versions changed).
4. Restart: `streamlit run canopy_configurator.py`.

---

## Important notes

- **Filenames are case- and space-sensitive.** Do not rename `Apec Imaging Logo.jpg` or `GEO Canopies logo.jpg` — the app looks for these exact names.
- **Pricing data is in `canopy_config.xlsx`.** The app caches the workbook with a 60-second TTL, so future edits to the Excel file propagate within ~1 minute with no restart. Code changes (`.py` file) still require a Streamlit restart.
- **Workbook has a hidden tab** (`Internal_PL`) protected by the workbook-structure password `cheap`. The app reads from it on startup; no action needed at deploy time.

---

## Quick verification after deploy

1. Open the live site.
2. The radio at the top reads "Is this quote for APEC Canopies or GEO Canopies?"
3. Selecting **APEC Canopies** → APEC logo displays.
4. Selecting **GEO Canopies** → GEO Canopies logo displays.
5. Fill in any test scenario and click **Calculate Canopy Price** — quote result should render.

If the header is missing or shows a broken image, confirm both `.jpg` files made it into the same directory as `canopy_configurator.py` with their exact filenames.
