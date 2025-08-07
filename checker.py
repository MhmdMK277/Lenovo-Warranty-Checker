import os
import csv
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

# Paths and folders
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
VERSIONS_DIR = os.path.join(BASE_DIR, "versions")
SERIAL_FILE = os.path.join(BASE_DIR, "serials.txt")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "warranty_results.csv")
LOGO_FILE = os.path.join(ASSETS_DIR, "PSEE.png")
CHROME_DRIVER_PATH = r"Path\\To\\Your\\Chromedriver.exe"  # Update this path

MAX_WORKERS = 5  # Number of concurrent threads
RETRY_COUNT = 2  # Number of retries on failure

# Ensure folders exist
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(VERSIONS_DIR, exist_ok=True)

# GUI elements will be created in main()
root: tk.Tk
status_label: tk.Label
progress_var: tk.DoubleVar


def prompt_headless() -> bool:
    """Ask the user if the script should run in headless mode."""
    return messagebox.askyesno(
        "Headless Mode",
        "Would you like to run in headless mode (no browser popups)? This may run faster.",
    )


def create_driver(headless: bool) -> webdriver.Chrome:
    """Create a configured Chrome driver."""
    options = Options()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    if headless:
        options.add_argument("--headless")
    service = Service(CHROME_DRIVER_PATH)
    return webdriver.Chrome(service=service, options=options)


def fetch_warranty(serial: str, headless: bool) -> tuple[str, str]:
    """Fetch warranty information for a serial number with retry logic."""
    for attempt in range(RETRY_COUNT + 1):
        try:
            driver = create_driver(headless)
            driver.get("https://pcsupport.lenovo.com/us/en/warranty-lookup")
            try:
                proceed_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(text(),'Proceed with United States of America')]")
                    )
                )
                proceed_button.click()
            except Exception:
                pass

            input_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "button-placeholder__input"))
            )
            input_box.clear()
            input_box.send_keys(serial)

            submit_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Submit')]") )
            )
            submit_button.click()

            try:
                end_date_element = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//div[@class='detail-property']//span[contains(text(), 'End Date')]/following-sibling::span")
                    )
                )
                end_date = end_date_element.text.strip() or "Not Found"
            except Exception:
                end_date = "Not Found"

            driver.quit()
            return serial, end_date
        except Exception:
            try:
                driver.quit()
            except Exception:
                pass
            if attempt < RETRY_COUNT:
                time.sleep(2)
            else:
                return serial, "Error"
    return serial, "Error"


def process_serials(headless: bool) -> None:
    """Process all serials using a thread pool and update the GUI progress bar."""
    if not os.path.exists(SERIAL_FILE):
        status_label.config(text="❌ serials.txt not found")
        return

    # Versioning of previous output
    if os.path.exists(OUTPUT_FILE):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        os.rename(OUTPUT_FILE, os.path.join(VERSIONS_DIR, f"warranty_results_{timestamp}.csv"))

    with open(SERIAL_FILE) as f:
        serials = [line.strip() for line in f if line.strip()]

    total = len(serials)
    status_label.config(text=f"Processing {total} serials...")
    root.update_idletasks()

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor, open(
        OUTPUT_FILE, "w", newline=""
    ) as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Serial", "EndDate"])
        futures = [executor.submit(fetch_warranty, s, headless) for s in serials]
        for idx, future in enumerate(as_completed(futures), start=1):
            serial, end_date = future.result()
            writer.writerow([serial, end_date])
            progress_var.set((idx / total) * 100)
            status_label.config(text=f"Processed {idx}/{total}")
            root.update_idletasks()

    status_label.config(text=f"✅ Done! Saved to {OUTPUT_FILE}")


def start_processing() -> None:
    headless = prompt_headless()
    threading.Thread(target=process_serials, args=(headless,), daemon=True).start()


def main() -> None:
    global root, status_label, progress_var
    root = tk.Tk()
    root.title("Lenovo Warranty Checker")

    try:
        logo = tk.PhotoImage(file=LOGO_FILE)
        tk.Label(root, image=logo).pack(pady=10)
        # Keep a reference to prevent garbage collection
        root.logo = logo
    except Exception:
        pass

    status_label = tk.Label(root, text="Initializing...", fg="blue")
    status_label.pack(pady=5)

    progress_var = tk.DoubleVar()
    ttk.Progressbar(root, variable=progress_var, maximum=100, length=300).pack(pady=10)

    root.after(100, start_processing)
    root.mainloop()


if __name__ == "__main__":
    main()
