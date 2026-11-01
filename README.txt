Account Confirmation Letter App
================================

FIRST TIME SETUP
1. Install Python from python.org (tick "Add Python to PATH" during install).
2. Open Command Prompt in this folder and run:
       pip install -r requirements.txt

START THE APP
- Windows: double-click start_app.bat
- Any system: run   streamlit run app.py   and open http://localhost:8501

FIRST RUN
The app opens on the Settings page. Enter the company name, address, phone, PAN,
default fiscal year and default "as at" text, then press "Save settings".
Nothing about the business is stored in the source code; it is saved locally in
data/settings.json.

PAGES
- Dashboard : totals, counts by nature, largest payables/receivables, recent letters.
- Data Entry: enter or edit a party. Choosing "Purchase and Sales" shows the Sales fields.
- Letters   : PDF per party (named after the party), view, edit, delete, download all.
- Settings  : company details, default fiscal year, stamp and signature.

STAMP AND SIGNATURE
Upload them once in Settings. White backgrounds are removed and empty margins are
trimmed automatically. Use the size and overlap sliders to adjust how they print.
They can be switched off on the Letters page for hand-signed copies.

NEW FISCAL YEAR
Change the default fiscal year and "as at" text in Settings. Older letters stay
available through the Fiscal year filter.

FILES (all local, all ignored by Git via .gitignore)
- data/parties.xlsx   all saved letters (can be opened in Excel; close it before saving in the app)
- data/settings.json  company settings
- assets/             stamp and signature images
Back up the data folder regularly.

NOTE
Write party names in English; Nepali characters will not print in the PDF.
