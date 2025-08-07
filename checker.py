import threading
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import time

chrome_driver_path = r"Path\To\Your\Chromedriver\exe\file"
serial_file = 'serials.txt'
output_file = 'warranty_results.xlsx'
logo_file = 'PSEE.png'

stop_flag = False
save_and_stop_flag = False
run_in_background_flag = False
results = []

def selenium_runner():
    global stop_flag, save_and_stop_flag, results, run_in_background_flag

    if os.path.exists(output_file):
        if messagebox.askyesno("File exists", f"{output_file} already exists.\nDelete it before starting?"):
            os.remove(output_file)
        else:
            update_status("❌ Stopped: existing file not deleted.")
            return

    with open(serial_file) as f:
        serials = [line.strip() for line in f if line.strip()]

    total_serials = len(serials)
    start_time = time.time()

    for idx, serial in enumerate(serials):
        if stop_flag:
            update_status("❌ Stopped without saving.")
            return

        if save_and_stop_flag:
            update_status("✅ Stopping and saving partial results.")
            break

        try:
            options = Options()
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            options.add_experimental_option('excludeSwitches', ['enable-logging'])

            service = Service(chrome_driver_path)
            driver = webdriver.Chrome(service=service, options=options)

            driver.get("https://pcsupport.lenovo.com/us/en/warranty-lookup")

            try:
                proceed_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Proceed with United States of America')]"))
                )
                proceed_button.click()
            except:
                pass

            input_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "button-placeholder__input"))
            )

            input_box.clear()
            input_box.send_keys(serial)

            submit_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Submit')]"))
            )
            submit_button.click()

            try:
                end_date_element = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, "//div[@class='detail-property']//span[contains(text(), 'End Date')]/following-sibling::span"))
                )
                end_date = end_date_element.text.strip()
            except:
                end_date = "Not found or expired"

            results.append({'Serial': serial, 'EndDate': end_date})

            driver.quit()
            service.stop()

            elapsed = time.time() - start_time
            avg_time = elapsed / (idx + 1)
            est_remaining = avg_time * (total_serials - idx - 1)

            if not run_in_background_flag:
                progress_var.set(((idx + 1) / total_serials) * 100)
                status_label.config(text=f"Processed {serial} ({idx + 1}/{total_serials}) - ~{int(est_remaining // 60)}m {int(est_remaining % 60)}s left")
                root.update_idletasks()

            pd.DataFrame(results).to_excel(output_file, index=False)
            time.sleep(3)

        except Exception as e:
            results.append({'Serial': serial, 'EndDate': f"Error: {str(e)}"})
            try:
                driver.quit()
                service.stop()
            except:
                pass
            if not run_in_background_flag:
                status_label.config(text=f"Error on {serial}")
                root.update_idletasks()
            pd.DataFrame(results).to_excel(output_file, index=False)
            time.sleep(3)
            continue

    if not stop_flag:
        pd.DataFrame(results).to_excel(output_file, index=False)
        if not run_in_background_flag:
            status_label.config(text=f"✅ Done! Excel saved as '{output_file}'.")

def stop_now():
    global stop_flag
    if messagebox.askyesno("Confirm", "Are you sure you want to stop immediately? You will NOT get the .xlsx file."):
        stop_flag = True
        if not run_in_background_flag:
            status_label.config(text="❌ Stopped immediately.")
            root.update_idletasks()

def stop_and_save():
    global save_and_stop_flag
    if messagebox.askyesno("Confirm", "Are you sure you want to stop and save what was processed so far?"):
        save_and_stop_flag = True
        if not run_in_background_flag:
            status_label.config(text="✅ Stopping and saving partial results...")
            root.update_idletasks()

def run_background():
    global run_in_background_flag
    if messagebox.askyesno("Background", "Run completely in the background? You won't see further updates, but it may run slightly faster."):
        run_in_background_flag = True
        root.withdraw()

def update_status(text):
    status_label.config(text=text)
    root.update_idletasks()

root = tk.Tk()
root.title("Lenovo Warranty Checker Control")

# Logo
try:
    logo_img = Image.open(logo_file)
    logo_img = logo_img.resize((150, 150), Image.Resampling.LANCZOS)
    logo_photo = ImageTk.PhotoImage(logo_img)
    logo_label = tk.Label(root, image=logo_photo)
    logo_label.pack(pady=10)
except:
    pass

stop_btn = tk.Button(root, text="❌ Stop Immediately", command=stop_now, bg="red", fg="white", width=25)
stop_btn.pack(pady=5)

save_stop_btn = tk.Button(root, text="✅ Stop and Save Partial", command=stop_and_save, bg="green", fg="white", width=25)
save_stop_btn.pack(pady=5)

bg_btn = tk.Button(root, text="▶️ Run in Background", command=run_background, bg="gray", fg="white", width=25)
bg_btn.pack(pady=5)

status_label = tk.Label(root, text="Status: Running...", fg="blue")
status_label.pack(pady=10)

progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=100, length=300)
progress_bar.pack(pady=10)

threading.Thread(target=selenium_runner, daemon=True).start()

root.mainloop()
