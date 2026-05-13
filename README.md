# Chapman Tools

Internal web apps for Chapman Automotive.

## Tools

- **Invoice Splitter** — Upload a combined invoice PDF, split into individual files, download as ZIP + Excel summary. Supports scanned PDFs via OCR.
- **Loaner Label Generator** — Upload a vehicle list (CSV or Excel), pick your start slot on an Avery 8460 sheet, download a print-ready PDF.

---

## Deploying to Streamlit Community Cloud

1. Push this folder to a GitHub repository (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub
3. Click **New app**
4. Select your repo, set branch to `main`, and set the main file to `Home.py`
5. Click **Deploy** — your app will be live in about 60 seconds

The `packages.txt` file tells Streamlit to install Tesseract OCR and Poppler automatically — no extra setup needed.

---

## Adding a New Tool

1. Create a new file in the `pages/` folder named `3_🔧_Tool_Name.py`
2. Add the tool description to the cards in `Home.py`
3. Push to GitHub — Streamlit redeploys automatically

---

## File Structure

```
chapman_tools/
├── Home.py                        ← Landing page / tool directory
├── requirements.txt               ← Python dependencies
├── packages.txt                   ← System dependencies (Tesseract, Poppler)
├── README.md
└── pages/
    ├── 1_📄_Invoice_Splitter.py   ← Invoice splitter tool
    └── 2_🏷️_Loaner_Labels.py     ← Loaner label generator
```

---

## Column Names Required

**Loaner Labels** — CSV or Excel must have these exact column headers:
`Unit Number`, `VIN`, `License Plate`, `Year`, `Make`, `Model`
