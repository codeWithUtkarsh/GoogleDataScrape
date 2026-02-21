# GoDaper — Web UI

A free Google Maps Data Finder. Search for any business type across UK postcodes.

![Python](https://img.shields.io/badge/Python-3.8+-blue)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

- **Web UI** — Clean dark-themed interface with real-time progress
- **Postcode lookup** — Enter any UK city and get all local postcodes (via free postcodes.io API)
- **Multi-postcode scraping** — Select one, many, or all postcodes to cover
- **Live progress** — Watch stores appear in real-time as they're scraped
- **Auto-deduplication** — Stores appearing in multiple postcode areas are only counted once
- **Excel export** — Formatted `.xlsx` with store data, postcode summary, and metadata
- **Phone numbers** — Extracts phone, address, rating, reviews, category, website, hours, coordinates
- **Upload & merge** — Upload an existing spreadsheet, change the query, and only new stores are appended
- **Cleanup** — One-click button to clear all cached files from the server

## Quick Start

```bash
# 1. Clone or copy the project
cd GoDaper

# 2. Install dependencies
pip install -r requirements.txt

# 3. Install browser (one-time)
playwright install chromium

# 4. Run the app
python app.py
```

Then open **http://localhost:5000** in your browser.

## How to Use

1. **Enter your search query** — e.g. "indian grocery store", "halal butcher", "pharmacy"
2. **Enter a UK city** — e.g. "Liverpool", "Manchester", "London"
3. **Click "Find Postcodes"** — fetches all postcode areas for that city
4. **Select postcodes** — click individual postcodes or "Select All"
5. **Click "Start Scraping"** — watch results stream in live
6. **Download Excel** — click the download button when complete


## Notes

- **UK postcodes only** — Uses postcodes.io (free, no key) for postcode lookup
- **Scraping speed** — Each postcode takes ~10-30 seconds depending on results
- **Full L1–L37 scan** — A complete Liverpool scan takes roughly 20-40 minutes
- **Be respectful** — The scraper includes delays to avoid overwhelming Google
- **Google may block** — If you scrape too aggressively, Google may show CAPTCHAs

## ⚠️ Disclaimer

This project is for **educational and personal use only**. Scraping Google Maps
may violate Google's [Terms of Service](https://policies.google.com/terms).
The author is not responsible for any misuse. Use at your own risk.