# Umra-Tracker

Small Node.js script that generates an Excel-based Umra payment tracker.

## What it generates
- `Umra.xlsx`
- One sheet per year (2026–2030) with:
  - A per-person table with monthly cells (data validation: `0` or `10` KM)
  - Year totals and conditional formatting (paid/unpaid coloring)
- A `PREGLED` (summary) sheet that aggregates totals across all years.

## Prerequisites
- Node.js installed

## Setup
1. Install dependencies:
   - `npm install`

## Run
From the project folder:
- `node create_umra_tracker.js`

After running, `Umra.xlsx` will be created/overwritten in the same folder.

## Dependencies
- `exceljs`

## Notes
- The script currently tracks 15 people (`Osoba 1` … `Osoba 15`) and uses a monthly amount of `10 KM`.