# Lenovo-Warranty-Checker

Batch-check Lenovo laptop warranties via serial numbers using Selenium.

## Features
- Optional headless mode prompt on startup
- Concurrent processing with progress bar
- Retry logic and error handling
- Outputs `output/warranty_results.csv` with versioned backups in `versions/`

## Usage
1. Place serial numbers in `serials.txt` (one per line).
2. Install ChromeDriver and update `CHROME_DRIVER_PATH` in `checker.py`.
3. Run `python checker.py`.

The script prompts for headless mode, displays progress in a GUI, and saves results for easy import into AssetIT.
