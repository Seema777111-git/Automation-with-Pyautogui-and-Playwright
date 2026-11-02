# ============================================================
# DAILY REPORT BOT
# ============================================================
# Desktop automation with PyAutoGUI:
#   1. Open Chrome on the CoinDesk Bitcoin page (Win+R)
#   2. Select and copy the page text (Ctrl+A, Ctrl+C)
#   3. Pull the live Bitcoin price out of the copied text
#   4. Ctrl+F the price so it is highlighted, screenshot as proof
#   5. Close Chrome (Alt+F4)
#   6. Write the daily Excel report
# ============================================================

import os
import re
import time
from datetime import datetime

import pyautogui
import pyperclip
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill


# ============================================================
# 1. SETTINGS
# ============================================================

pyautogui.FAILSAFE = True      # move the mouse to a screen corner to stop
pyautogui.PAUSE = 0.3

URL = "https://www.coindesk.com/price/bitcoin/"

# Chrome profile to open (skips the "Who's using Chrome?" picker).
# See chrome://version -> "Profile Path" for the folder name, e.g. "Default" or "Profile 1".
CHROME_PROFILE = "Profile 1"

PAGE_LOAD_SECONDS = 10

# Save reports next to this script (works on any PC)
OUTPUT_FOLDER = os.path.dirname(os.path.abspath(__file__))

today = datetime.now()
date_time = today.strftime("%Y-%m-%d %H:%M:%S")
date_name = today.strftime("%Y-%m-%d")

excel_file = os.path.join(OUTPUT_FOLDER, f"daily_report_{date_name}.xlsx")
screenshot_file = os.path.join(OUTPUT_FOLDER, f"daily_report_{date_name}.png")

SOURCE = "CoinDesk"


# ============================================================
# 2. OPEN CHROME ON THE WEBSITE
# ============================================================

def open_website():

    print("STEP 1: Opening Chrome on CoinDesk...")

    # Win+R opens the Run box; launching Chrome with a profile and the URL
    # avoids the profile picker and the address-bar typing.
    pyautogui.hotkey("win", "r")
    time.sleep(1)

    command = (
        f'chrome --profile-directory="{CHROME_PROFILE}" '
        f'--new-window --start-maximized {URL}'
    )
    pyautogui.write(command, interval=0.02)
    pyautogui.press("enter")

    print(f"Waiting {PAGE_LOAD_SECONDS} seconds for the page to load...")
    time.sleep(PAGE_LOAD_SECONDS)


# ============================================================
# 3. COPY THE PAGE AND FIND THE PRICE
# ============================================================

def find_price(text):
    """Return the Bitcoin price from the copied page text."""

    # On the page the price follows the ticker: "BTC  #1  $77,383.11"
    match = re.search(r"BTC\s*(?:#\d+\s*)?(\$[\d,]+\.\d{2})", text)

    if match:
        return match.group(1)

    # Fallback: first dollar amount with cents
    match = re.search(r"\$[\d,]+\.\d{2}", text)

    return match.group(0) if match else None


def get_web_data():

    print("STEP 2: Copying the page and finding the Bitcoin price...")

    price = None

    for attempt in range(1, 4):

        pyperclip.copy("")

        # Select all text on the page and copy it
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)
        pyautogui.hotkey("ctrl", "c")
        time.sleep(1.5)

        # Clear the selection highlight
        pyautogui.press("esc")
        pyautogui.hotkey("ctrl", "home")

        price = find_price(pyperclip.paste())

        if price:
            break

        print(f"  Price not on the page yet (attempt {attempt}), waiting...")
        time.sleep(4)

    data = price or "PRICE NOT FOUND"

    print()
    print("BITCOIN PRICE:", data)
    print()

    return data


# ============================================================
# 4. SCREENSHOT OF THE PAGE
# ============================================================

def take_screenshot(data):

    print("STEP 3: Highlighting the price and taking a screenshot...")

    # Ctrl+F the captured price: Chrome scrolls to it and highlights it,
    # so the screenshot shows exactly the value that was read.
    if data != "PRICE NOT FOUND":
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.8)
        pyautogui.write(data.lstrip("$"), interval=0.05)
        time.sleep(1.5)

    pyautogui.screenshot().save(screenshot_file)

    # Close the find bar
    pyautogui.press("esc")

    print("Screenshot saved:", screenshot_file)


# ============================================================
# 5. CLOSE CHROME
# ============================================================

def close_chrome():

    print("STEP 4: Closing Chrome...")

    pyautogui.hotkey("alt", "f4")
    time.sleep(1.5)


# ============================================================
# 6. WRITE THE EXCEL REPORT
# ============================================================

def save_excel(data):

    print("STEP 5: Writing the Excel report...")

    headers = ["Date & Time", "Bitcoin Price (USD)", "Source", "Comment"]

    comment = (
        "Market information captured successfully"
        if data != "PRICE NOT FOUND"
        else "Price could not be read from the page"
    )

    # Add a row to today's report (several runs a day -> several rows)
    if os.path.exists(excel_file):
        workbook = load_workbook(excel_file)
        sheet = workbook.active
    else:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Daily Report"
        sheet.append(headers)

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")

    sheet.append([date_time, data, SOURCE, comment])

    for column, width in zip("ABCD", [22, 22, 14, 44]):
        sheet.column_dimensions[column].width = width

    workbook.save(excel_file)

    print("Excel saved:", excel_file)


# ============================================================
# 7. MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print("          DAILY REPORT AUTOMATION")
    print("==============================================")
    print()

    open_website()
    data = get_web_data()
    take_screenshot(data)
    close_chrome()
    save_excel(data)

    print()
    print("==============================================")
    print("             REPORT COMPLETED")
    print("==============================================")
    print()
    print("Bitcoin price:", data)
    print("Excel        :", excel_file)
    print("Screenshot   :", screenshot_file)
    print()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()
