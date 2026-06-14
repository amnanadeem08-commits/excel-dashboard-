# Excel Dashboard Studio

Excel Dashboard Studio is a local-first Streamlit app for generating editable, sellable Excel dashboard templates. It creates branded `.xlsx` workbooks with business details, optional logo upload, selected color palettes, sample data, formulas, instructions, and marketplace copy.

## Features

- Streamlit UI for beginner-friendly workbook generation
- Five Excel template options
- Six built-in color palettes
- Business details, currency, and Tax/VAT/GST fields
- Logo upload with automatic workbook insertion
- No-logo placeholder when a logo is not provided
- Formula-driven workbook sheets with lookup/reference logic
- Data validation dropdowns and conditional formatting where useful
- Printable dashboard, invoice, or receipt areas
- Local output folder for generated workbooks
- Rule-based marketplace copy generation with no external API

## Installation

1. Open a terminal in the project folder:

```powershell
cd D:\exceldashboard
```

2. Install dependencies:

```powershell
pip install -r requirements.txt
```

## Run Command

```powershell
streamlit run app.py
```

Then open the local Streamlit URL shown in the terminal, usually:

```text
http://localhost:8501
```

## Template List

- Small Business Bookkeeping / Profit & Loss Dashboard
- Sales Dashboard
- Inventory Management Dashboard
- Invoice & Receipt Template with Logo
- CRM Customer Tracker

## Output Folder

Generated Excel workbooks are saved locally in:

```text
outputs/
```

Marketplace copy files are saved locally in:

```text
outputs/marketplace_copy/
```

Each generated workbook can also be downloaded directly from the Streamlit app using the `Download .xlsx` button.

## Notes for Marketplace Selling

- Review each generated workbook before listing it for sale.
- Replace or polish screenshots using the final workbook design and selected palette.
- Use the generated marketplace copy as a starting point for Etsy, Fiverr, Gumroad, or other platforms.
- Keep product descriptions honest: mention that the workbook is editable and includes sample data.
- Include screenshots of the dashboard, input sheets, lookup sheets, summary sheets, and instructions.
- Test files in Microsoft Excel before uploading to a marketplace.
- Do not include private business data, secrets, API keys, or customer information in sellable files.

## Local-First Notes

This project does not use API keys, paid plugins, cloud services, or external data connections. Workbook generation runs locally with Python, Streamlit, openpyxl, and Pillow.
