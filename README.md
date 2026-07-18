# Project Overview

The **Canopy Configurator** is a web-based quoting and pricing tool designed to configure canopy installations (for gas and diesel stations) and imaging (branding) services. 

* **Who it is for**: Sales representatives and internal staff.
* **Main features**: 
  * Interactive UI for single canopy, double canopy (gas + diesel), and imaging-only quoting.
  * Dynamic pricing engine driven by an external Excel configuration file.
  * Automated generation of formal customer proposals in Word format (`.docx`).
  * Automated generation of internal profitability reports in Excel format (`.xlsx`).
* **Business purpose**: Streamlines the quoting process, ensures consistent pricing based on configurable margins and costs, and provides an instant turnaround for both customer-facing proposals and internal financial reviews.

# Tech Stack

* **Python 3**: Core programming language.
* **Streamlit**: Web application framework used to build the interactive UI quickly.
* **python-docx**: Used to programmatically generate formatted `.docx` proposal documents.
* **openpyxl**: Used to read pricing configuration data from an Excel file and to generate `.xlsx` profitability reports.

# Getting Started

### Prerequisites
* Python 3.9+

### Installation
1. Clone the repository to your local machine.
2. Create a virtual environment:
   ```bash
   python -m venv .venv
   ```
3. Activate the virtual environment:
   * **Windows**: `.venv\Scripts\activate`
   * **Mac/Linux**: `source .venv/bin/activate`
4. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### Environment variables
No environment variables are required for local development.

### Running locally
To run the Streamlit app locally, execute:
```bash
streamlit run canopy_configurator.py
```
Alternatively, you can use the startup script: `bash start.sh`

### Build & Deployment commands
* **Build**: No build step is required for Streamlit applications.
* **Deployment**: Deployments are handled automatically via Streamlit Community Cloud upon pushing to the connected Git branch:
  ```bash
  git push origin main
  ```

# Project Structure

* `canopy_configurator.py` → The main Streamlit application and UI logic (entry point).
* `proposal_writer.py` → Logic to generate Word (`.docx`) proposals based on quote data.
* `profitability_report.py` → Logic to generate Excel (`.xlsx`) profitability reports showing cost, retail, and profit breakdowns.
* `canopy_config.xlsx` → The central pricing configuration file containing all material costs, labor rates, and margins. It is read dynamically at startup.
* `requirements.txt` → Python dependencies.
* `start.sh` & `Procfile` → Deployment scripts (Streamlit Cloud / Railway).

# Database

* **Database provider**: This project does not use a traditional relational database.
* **Main models/tables**: Configuration and pricing rates are stored in the `canopy_config.xlsx` file. Different tabs serve as tables (e.g., `Internal_PL`, `Material_Rates`, `MISC_Tiers`, `Labor_Days`, etc.).
* **Important relationships**: The Streamlit app reads this file into memory on startup and maps the data to calculation variables. Quote tracking is not persistent; data is processed in memory per session.

# Authentication

* **How authentication works**: The main quoting tool is accessible to all sales reps. However, specific sensitive actions (like viewing the profitability report or overriding minimum GP margins) are gated by hardcoded passwords.
* **User roles and permissions**: "Admin" or internal leadership roles can access the Profitability Report section at the bottom of the configurator by entering the correct password (currently `showmethemoney`).

# Key Features

### Pricing Engine & Configurator
* **Purpose**: Collects project requirements (dimensions, dispensers, branding, location) and calculates costs and retail prices.
* **Important files**: `canopy_configurator.py`, `canopy_config.xlsx`
* **How it works**: Uses inputs from the Streamlit UI, matches them with rates from `canopy_config.xlsx`, and applies defined gross profit margins to arrive at the final retail price.

### Proposal Generation
* **Purpose**: Creates a formatted Word document proposal to be sent to the customer.
* **Important files**: `proposal_writer.py`
* **How it works**: Uses `python-docx` to programmatically build the document with tables, headers, and legal terms based on the active quote data.

### Profitability Report
* **Purpose**: Provides internal leadership with a line-by-line breakdown of costs vs. retail pricing.
* **Important files**: `profitability_report.py`
* **How it works**: Builds an `.xlsx` file using `openpyxl`. It embeds Excel formulas so that if costs or retail values are manually adjusted later by leadership, the margins recalculate automatically.

# Third-Party Services

* **Streamlit Community Cloud**: Hosting provider for the application (`www.alihusain.me`).

# Deployment

* **Hosting provider**: Streamlit Community Cloud (formerly integrated with Railway).
* **Environment setup**: The GitHub repository is connected directly to Streamlit Cloud.
* **Production deployment process**: Changes are deployed automatically when code is pushed to the connected branch. Pricing changes can be made simply by updating the `canopy_config.xlsx` file and pushing the change, requiring no Python code updates.

# Common Tasks

### Update Pricing or Labor Rates
Edit the relevant tab in `canopy_config.xlsx`, save, and push the file to the repository. The application will use the new values on the next load. Do not rename tabs or columns.

### Add a UI input field
Modify `canopy_configurator.py` using Streamlit widgets (e.g., `st.number_input`, `st.selectbox`). Ensure the captured value is added to the quote dictionary (`q`) and passed to downstream functions if needed.

### Modify Proposal Format
Edit `proposal_writer.py` to add or change document paragraphs, tables, or styling. Use the `python-docx` library functions (like `doc.add_paragraph()`).

### Deploy changes
Commit your changes to Git and push to the main branch. Streamlit Cloud will auto-redeploy the application in approximately 1–2 minutes.

# Known Issues & Technical Debt

* **No persistent storage**: Quotes are not currently saved to a database (the Google Sheets tracker was removed in v14). If a quote document is lost, the inputs must be re-entered manually.
* **Hardcoded passwords**: The profitability report password (`showmethemoney`) and the GP override password (`cheap`) are hardcoded directly in `canopy_configurator.py`.
* **File size and complexity**: The application logic in `canopy_configurator.py` is quite large (~1500 lines). It mixes UI layout, state management, and business logic.

# Notes for Future Developers

* **Important design decisions**: All pricing variables, margins, and material costs are deliberately externalized to `canopy_config.xlsx`. This allows business users to update pricing without developer intervention. **Do not hardcode pricing values in Python.**
* **Things to be careful with**: Tab names and column structures in `canopy_config.xlsx` must remain exactly as expected by the code. Changing a tab name or structure will break the data loader.
* **Areas that need refactoring**: Break down `canopy_configurator.py` into smaller modular components (e.g., separate files for layout, calculation logic, and Streamlit state management).
* **Business rules that are not obvious**: 
  * Double canopy quotes apply a Gross Profit (GP) reduction (currently 3%) to the second canopy.
  * Sales tax rounding is specifically implemented to round to cents *before* the total is summed to ensure the Word proposal and Excel profitability report tie out perfectly to the penny.
