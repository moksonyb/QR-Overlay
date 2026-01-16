# QR Bulk Adder (Web)

A Flask web app that overlays QR codes from a CSV column onto a PDF, one QR per page.

- CSV must have a header row.
- CSV rows (excluding header) must match the number of PDF pages and be in the correct order.
- QR size and position are in PDF points (1 pt = 1/72 inch).
- Origin is bottom-left, same as PDF coordinate system.
- Includes a layout-only preview (page outline and QR box) in the browser.

## Setup

```sh
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## Run

```sh
python app.py
```

Open http://127.0.0.1:5000 and:
- Select the PDF and CSV files.
- Choose the CSV column from the dropdown (populated from CSV header).
- Adjust QR size and X/Y positions.
- Click Preview to visualize placement.
- Click Generate to download the resulting PDF.

## CSV Format
First row is treated as a header. Example:

```
ID,Name,Email
INV-0001,John,john@example.com
INV-0002,Jane,jane@example.com
INV-0003,Bob,bob@example.com
```

Select "ID" column to generate QR codes from INV-0001, INV-0002, INV-0003.