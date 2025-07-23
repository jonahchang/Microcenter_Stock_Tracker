import re
import time
import os
import tkinter as tk
import openpyxl
from tkinter import Tk, filedialog, simpledialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side

# List of product URLs to check
product_urls = [
    # PSU PRODUCTS

    # ROG-THOR-1000P2-GAMING
    "https://www.microcenter.com/product/666611/asus-rog-thor-1000-watt-80-plus-platinum-atx-fully-modular-power-supply",
    # ROG-STRIX-1200P-GAMING
    "https://www.microcenter.com/product/695232/asus-rog-strix-1200-watt-80-plus-platinum-atx-fully-modular-power-supply-atx-31-compatible",
    # ROG-LOKI-1000P-SFX-L-GAMING
    "https://www.microcenter.com/product/664884/asus-rog-loki-1000-watt-80-plus-platinum-sfx-l-fully-modular-power-supply-black-atx-30-compatible",
    # ROG-LOKI-850P-WHITE-SFX-L-GAMING
    "https://www.microcenter.com/product/664885/asus-rog-loki-850-watt-80-plus-platinum-sfx-l-fully-modular-power-supply-white-atx-30-compatible",
    # ROG-LOKI-850P-SFX-L-GAMING
    "https://www.microcenter.com/product/664883/asus-rog-loki-850-watt-80-plus-gold-sfx-l-fully-modular-power-supply-black-atx-30-compatible",
    # ROG-STRIX-850G-AURA-GAMING
    "https://www.microcenter.com/product/669273/asus-rog-strix-gold-aura-edition-850-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    # ROG-STRIX-750G-AURA-GAMING
    "https://www.microcenter.com/product/669274/asus-rog-strix-gold-aura-edition-750-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    # TUF-GAMING-1200G
    "https://www.microcenter.com/product/676964/asus-tuf-gaming-1200-watt-80-plus-gold-atx-fully-modular-power-supply",
    # TUF-GAMING-1000G
    "https://www.microcenter.com/product/665308/asus-tuf-gaming-1000-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    #TUF-GAMING-850G
    "https://www.microcenter.com/product/665319/asus-tuf-gaming-850-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    # TUF-GAMING-750G
    "https://www.microcenter.com/product/665320/asus-tuf-gaming-750-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    # AP-850G
    "https://www.microcenter.com/product/675843/asus-prime-850-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    # AP-750G
    "https://www.microcenter.com/product/675842/asus-prime-750-watt-80-plus-gold-atx-fully-modular-power-supply-atx-30-compatible",
    
    # AIO PRODUCTS
    
    # ROG RYUJIN III 360 ARGB EXTREME WHT
    "https://www.microcenter.com/product/690065/asus-asus-rog-ryuyjin-iii-360-argb-extreme-360mm-all-in-one-liquid-cpu-cooling-kit-white",
    # ROG RYUJIN III 360 ARGB EXTREME
    "https://www.microcenter.com/product/690066/asus-rog-ryujin-iii-argb-extreme-360mm-all-in-one-liquid-cpu-cooling-kit-black",
    # ROG RYUJIN III 360
    "https://www.microcenter.com/product/668461/asus-rog-ryujin-iii-360mm-all-in-one-liquid-cpu-cooling-kit",
    # ProArt LC 420
    "https://www.microcenter.com/product/678856/asus-proart-lc-420mm-all-in-one-liquid-cpu-cooling-kit-black",
    
    # CHASSIS PRODUCTS
    
    # GR701 ROG HYPERION
    "https://www.microcenter.com/product/664435/asus-asus-rog-hyperion-gr701-tempered-glass-eatx-full-tower-computer-case-black",
    # GX601 ROG STRIX HELIOS CASE/WT/AL/WITH HANDLE
    "https://www.microcenter.com/product/625183/asus-rog-strix-helios-gx601-rgb-tempered-glass-atx-mid-tower-computer-case-white-edition",
    # GX601 ROG STRIX HELIOS CASE/BK/AL/WITH HANDLE
    "https://www.microcenter.com/product/609942/asus-rog-strix-helios-gx601-rgb-tempered-glass-atx-mid-tower-computer-case-black",
    # PA602 ProArt Case
    "https://www.microcenter.com/product/676302/asus-proart-pa602-tempered-glass-eatx-mid-tower-computer-case-black",
    # PROART PA401 WOOD TG PWM BLACK
    "https://www.microcenter.com/product/690056/asus-proart-pa401-wood-edition-tempered-glass-atx-mid-tower-computer-case-black",
    # GT502 TUF GAMING CASE/BLK
    "https://www.microcenter.com/product/662252/asus-tuf-gaming-gt502-tempered-glass-atx-mid-tower-computer-case-black",
    # GT502 TUF GAMING CASE/WHT
    "https://www.microcenter.com/product/662254/asus-tuf-gaming-gt502-tempered-glass-atx-mid-tower-computer-case-white",
    # GT501 TUF GAMING CASE/GRY/WITH HANDLE
    "https://www.microcenter.com/product/601243/asus-tuf-gaming-gt501-rgb-tempered-glass-atx-mid-tower-computer-case",
    # TUF GAMING GT302 ARGB BLACK
    "https://www.microcenter.com/product/679946/asus-tuf-gaming-gt302-argb-tempered-glass-atx-mid-tower-computer-case-black",
    # TUF GAMING GT302 ARGB  WHT
    "https://www.microcenter.com/product/679945/asus-tuf-gaming-gt302-argb-tempered-glass-atx-mid-tower-computer-case-white",
    # A31 PLUS/BK/TG/ARGB// 
    "https://www.microcenter.com/product/690543/asus-a31-plus-tempered-glass-atx-mid-tower-computer-case-black",
    # AP201 ASUS PRIME CASE MESH
    "https://www.microcenter.com/product/651914/asus-prime-ap201-microatx-mini-tower-computer-case-black",
    # AP201 ASUS PRIME CASE MESH WHITE EDITION
    "https://www.microcenter.com/product/651917/asus-prime-ap201-microatx-mini-tower-computer-case-white",
    # Add more URLs as needed
]

# Mapping of store ID to store name (must match Excel header format)
store_map = {
    "195": "Santa Clara", "101": "Tustin", "181": "Denver", "185": "Miami", "065": "Duluth", "041": "Marietta", 
    "151": "Chicago", "025": "Westmont", "165": "Indianapolis", "191": "Overland Park", "121": "Cambridge", 
    "085": "Rockville", "125": "Parkville", "055": "Madison Heights", "045": "St Louis Park", "095": "Brentwood", 
    "175": "Charlotte", "075": "New Jersey", "171": "Westbury", "115": "Brooklyn", "145": "Flushing", "105": "Yonkers", 
    "141": "Colombus", "051": "Mayfield Heights", "071": "Sharonville", "061": "St Davids", "155": "Houston",
    "131": "Dallas", "081": "Fairfax"
}

# Model names
model_names = [
    "ROG-THOR-1000P2-GAMING",
    "ROG-STRIX-1200P-GAMING",
    "ROG-LOKI-1000P-SFX-L-GAMING",
    "ROG-LOKI-850P-WHITE-SFX-L-GAMING",
    "ROG-LOKI-850P-SFX-L-GAMING",
    "ROG-STRIX-850G-AURA-GAMING",
    "ROG-STRIX-750G-AURA-GAMING",
    "TUF-GAMING-1200G",
    "TUF-GAMING-1000G",
    "TUF-GAMING-850G",
    "TUF-GAMING-750G",
    "AP-850G",
    "AP-750G",
    "ROG RYUJIN III 360 ARGB EXTREME WHT",
    "ROG RYUJIN III 360 ARGB EXTREME",
    "ROG RYUJIN III 360",
    "ProArt LC 420",
    "GR701 ROG HYPERION",
    "GX601 ROG STRIX HELIOS CASE/WT/AL/WITH HANDLE",
    "GX601 ROG STRIX HELIOS CASE/BK/AL/WITH HANDLE",
    "PA602 ProArt Case",
    "PROART PA401 WOOD TG PWM BLACK",
    "GT502 TUF GAMING CASE/BLK",
    "GT502 TUF GAMING CASE/WHT",
    "GT501 TUF GAMING CASE/GRY/WITH HANDLE",
    "TUF GAMING GT302 ARGB BLACK",
    "TUF GAMING GT302 ARGB  WHT",
    "A31 PLUS/BK/TG/ARGB// ",
    "AP201 ASUS PRIME CASE MESH",
    "AP201 ASUS PRIME CASE MESH WHITE EDITION"
]

# Stock highlighting colors
green_fill = PatternFill(start_color='00FF00', end_color='00FF00',  fill_type='solid')
yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00',  fill_type='solid')
blue_fill = PatternFill(start_color='83CCEB', end_color='83CCEB',  fill_type='solid')

def normalize_cell_value(value):
    if value is None:
        return 0
    val = str(value).strip().upper()
    if val == 'NO STOCK':
        return 0
    if val == 'N/A':
        return 'N/A'
    if val == '25+':
        return 25
    try:
        int_val = int(val)
        return 25 if int_val >= 25 else int_val
    except ValueError:
        return 0
    

def analyze_stock(filename):
    wb = openpyxl.load_workbook(filename)
    sheetnames = wb.sheetnames
    
    for i in range(2, len(sheetnames)):
        prev_sheet = wb[sheetnames[i - 1]]
        curr_sheet = wb[sheetnames[i]]
        for row in curr_sheet.iter_rows():
            for cell in row:
                prev_val = normalize_cell_value(prev_sheet.cell(row=cell.row, column=cell.column).value)
                curr_val = normalize_cell_value(cell.value)

                if (prev_val == 'N/A' and isinstance(curr_val, int)) or (curr_val == 'N/A' and isinstance(prev_val, int)):
                    cell.fill = blue_fill
                elif isinstance(prev_val, int) and isinstance(curr_val, int):
                    if curr_val > prev_val:
                        cell.fill = green_fill
                    elif curr_val < prev_val:
                        cell.fill = yellow_fill

    wb.save(filename)

# Format the sheet
def format_new_sheet(ws):
    for i, name in enumerate(model_names, start=2):
        ws[f"B{i}"] = name

    ws.merge_cells("A2:A14")
    ws["A2"] = "Power Supply"
    ws.merge_cells("A15:A18")
    ws["A15"] = "AIO Liquid CPU Cooler"
    ws.merge_cells("A19:A31")
    ws["A19"] = "Chassis"

    ws["AG1"] = "INDIVIDUAL TOTALS"
    ws["AH1"] = "CATEGORY TOTALS"
    ws["AH2"] = "PSU"
    ws["AH15"] = "AIO"
    ws["AH19"] = "CHASSIS"

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws["A1:AE31"]:
        for cell in row:
            cell.border = border

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[col_letter].width = adjusted_width


def product_sums(ws):
    for row in range(2, 32):
        ws[f"AG{row}"] = f"=SUM(C{row}:AE{row})"

    ws["AH3"] = "=SUM(AG2:AG14)"
    ws["AH16"] = "=SUM(AG15:AG18)"
    ws["AH20"] = "=SUM(AG19:AG31)"

def run_stock_tracker(target_wb, sheet_name):
    # Setup Selenium driver
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    # Setup worksheet
    ws = target_wb.create_sheet(title=sheet_name)
    headers = ["Product Category", "Model"] + list(store_map.values())
    ws.append(headers)

    # Start scanning for URLs
    for url in product_urls:
        driver.get(url)
        time.sleep(1)
        product_name = driver.title.split("-")[0].strip()
        row = [None]
        row.append(product_name)
        print(f"\nChecking stock for: {product_name}")

        for store_id, store_name in store_map.items():
            try:
                stock = get_stock(url, store_id, driver)
            except:
                stock = 0
            print(f"{store_name}: {stock}")
            row.append(stock)
        ws.append(row)

    format_new_sheet(ws)
    driver.quit()
    
def get_stock(url, store_id, driver):
    driver.get("https://www.microcenter.com")
    time.sleep(0.5)
    driver.add_cookie({
        'name': 'storeSelected', 'value': store_id, 'domain': '.microcenter.com', 'path': '/',
        'secure': True, 'httpOnly': False
    })
    driver.get(url)
    time.sleep(1)
    try:
        stock_element = driver.find_element(By.CSS_SELECTOR, "#pnlInventory > div > div > span > span.inventoryCnt")
        stock_text = stock_element.text.strip()
        if "25+" in stock_text:
            return "25+"
        else:
            match = re.search(r"\d+", stock_text)
            return int(match.group(0)) if match else 0
    except:
        return 0

def main():
    # Prompt user for file save location
    root = Tk()
    root.withdraw()  # Hide the main window

    # Ask the user for the week number (default is 1)
    week_number = simpledialog.askstring("Week Number", "Enter the week number (e.g., 9, 10, 11, etc):")
    if not week_number:
        week_number = 1
    
    use_existing = messagebox.askyesno("Stock Tracker", "Would you like to edit an existing Excel file?")

    # If the user owns the stock tracker file
    if use_existing:
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            messagebox.showinfo("Cancelled", "No file selected. Exiting.")
            return
        
        # Run the stock tracker and coloring
        wb = load_workbook(file_path)
        sheet_name = f"WK{week_number}"
        run_stock_tracker(wb, sheet_name)
        wb.save(file_path)
        analyze_stock(file_path)
        messagebox.showinfo("Done", f"Stock data added and highlighted in:\n{file_path}")
    else:
        # If the user opts to create an independent sheet with this week's stock
        wb = Workbook()
        ws = wb.active
        sheet_name = f"WK{week_number}"
        ws.title = sheet_name
        run_stock_tracker(wb, sheet_name)
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save new stock report as")
        if save_path:
            wb.save(save_path)
            messagebox.showinfo("Saved", f"New stock report saved as:\n{save_path}")
        else:
            messagebox.showinfo("Cancelled", "No save location selected.")

if __name__ == "__main__":
    main()