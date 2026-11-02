# ============================================================
# DAILY REPORT BOT
# ============================================================

import os
import time
import re
from datetime import datetime

import pyautogui
import pyperclip


# ============================================================
# 1. SETTINGS
# ============================================================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

URL = "https://www.coindesk.com/price/bitcoin/"

# Save reports next to this script (works on any PC)
OUTPUT_FOLDER = os.path.dirname(os.path.abspath(__file__))

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

today = datetime.now()

date_time = today.strftime("%Y-%m-%d %H:%M:%S")
date_name = today.strftime("%Y-%m-%d")

excel_filename = f"daily_report_{date_name}.xlsx"
screenshot_filename = f"daily_report_{date_name}.png"

excel_file = os.path.join(
    OUTPUT_FOLDER,
    excel_filename
)

screenshot_file = os.path.join(
    OUTPUT_FOLDER,
    screenshot_filename
)

COMMENT = "Market information captured successfully"


# ============================================================
# 2. OPEN CHROME
# ============================================================

def open_chrome():

    print("STEP 1: Opening Chrome...")

    pyautogui.hotkey("win", "r")
    time.sleep(0.6)

    pyautogui.write(
        "chrome",
        interval=0.03
    )

    pyautogui.press("enter")

    time.sleep(4)


# ============================================================
# 3. OPEN WEBSITE
# ============================================================

def open_website():

    print("STEP 2: Opening website...")

    pyautogui.hotkey("ctrl", "t")
    time.sleep(0.5)

    pyautogui.hotkey(
        "ctrl",
        "l"
    )
    time.sleep(0.3)

    pyperclip.copy(URL)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)

    pyautogui.press("enter")

    time.sleep(4)

    pyautogui.press("f5")

    # Wait for the page to finish loading before searching it
    # (searching too early finds nothing -> "PRICE NOT FOUND")
    time.sleep(8)


# ============================================================
# 4. FIND BITCOIN, SELECT PRICE AND COPY
# ============================================================

def get_web_data():

    print("STEP 3: Finding Bitcoin price...")

    # --------------------------------------------------------
    # Find Bitcoin
    # --------------------------------------------------------

    pyautogui.hotkey("ctrl", "f")
    time.sleep(1)

    pyautogui.write(
        "Bitcoin",
        interval=0.05
    )

    # Go to the required Bitcoin occurrence
    pyautogui.press("enter")
    time.sleep(1)

    # Close Find bar
    pyautogui.press("esc")
    time.sleep(1)

    # --------------------------------------------------------
    # SELECT BITCOIN INFORMATION UP TO PRICE
    # --------------------------------------------------------

    pyautogui.keyDown("shift")

    pyautogui.press(
        "down",
        presses=3,
        interval=0.2
    )

    pyautogui.press("end")

    pyautogui.keyUp("shift")

    time.sleep(1)

    # --------------------------------------------------------
    # COPY CURRENT SELECTION
    # --------------------------------------------------------

    pyautogui.hotkey(
        "ctrl",
        "c"
    )

    time.sleep(3)

    # Read clipboard
    raw_data = pyperclip.paste().strip()

    print()
    print("COPIED DATA:")
    print(repr(raw_data))
    print()

    # --------------------------------------------------------
    # GET ONLY THE PRICE
    # --------------------------------------------------------

    price_match = re.search(
        r"\$[\d,]+(?:\.\d+)?",
        raw_data
    )

    if price_match:

        data = price_match.group(0)

    else:

        data = "PRICE NOT FOUND"

    print()
    print("PRICE TO EXCEL:")
    print(data)
    print()

    # --------------------------------------------------------
    # CLOSE CHROME
    # --------------------------------------------------------

    print("Closing Chrome...")

    pyautogui.hotkey(
        "alt",
        "f4"
    )

    time.sleep(2)

    return data
	
# ============================================================
# 5. OPEN EXCEL
# ============================================================

def open_excel():

    print("STEP 4: Opening Excel...")

    pyautogui.hotkey(
        "win",
        "r"
    )

    time.sleep(0.6)

    pyautogui.write(
        "excel",
        interval=0.03
    )

    pyautogui.press("enter")

    time.sleep(4)


# ============================================================
# 6. ENTER DATA INTO EXCEL
# ============================================================

def enter_excel_data(data):

    print("STEP 5: Entering data into Excel...")

    pyautogui.hotkey(
        "ctrl",
        "n"
    )

    time.sleep(1.5)

    pyautogui.hotkey(
        "ctrl",
        "home"
    )

    # Headers

    pyautogui.write(
        "Date & Time",
        interval=0.02
    )

    pyautogui.press("tab")

    pyautogui.write(
        "Fetched Data",
        interval=0.02
    )

    pyautogui.press("tab")

    pyautogui.write(
        "Comment",
        interval=0.02
    )

    pyautogui.press("enter")

    # Timestamp

    pyautogui.write(
        date_time,
        interval=0.02
    )

    pyautogui.press("tab")

    # Website data

    pyperclip.copy(data)

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    pyautogui.press("tab")

    # Comment

    pyautogui.write(
        COMMENT,
        interval=0.02
    )

    pyautogui.press("enter")

    time.sleep(1)

    print("Data entered successfully.")


# ============================================================
# 7. SAVE EXCEL - EXACT LOCATION
# ============================================================

def save_excel():

    print("STEP 7: Saving Excel...")

    pyautogui.press("f12")

    time.sleep(2)

    pyautogui.hotkey(
        "ctrl",
        "a"
    )

    time.sleep(0.5)

    pyperclip.copy(excel_file)

    pyautogui.hotkey(
        "ctrl",
        "v"
    )

    time.sleep(0.5)

    pyautogui.press("enter")

    time.sleep(2)


    # ========================================================
    # REPLACE EXISTING FILE -> YES
    # ========================================================

    print("Checking for Replace confirmation...")

    # Excel's Replace dialog supports Alt+Y for Yes.
    pyautogui.hotkey(
        "alt",
        "y"
    )

    time.sleep(2)


    # ========================================================
    # KEEP CURRENT FORMAT
    # ========================================================

    pyautogui.hotkey(
        "alt",
        "y"
    )

    time.sleep(2)


    # ========================================================
    # FINAL SAVE
    # ========================================================

    pyautogui.hotkey(
        "ctrl",
        "s"
    )

    time.sleep(2)


    # ========================================================
    # VERIFY FILE
    # ========================================================

    if os.path.exists(excel_file):

        print()
        print("EXCEL SAVED SUCCESSFULLY")
        print(excel_file)
        print()

    else:

        print()
        print("ERROR: Excel file was NOT found.")
        print("Expected location:")
        print(excel_file)
        print()


# ============================================================
# 8. TAKE SCREENSHOT OF EXCEL
# ============================================================

def take_screenshot():

    print("STEP 6: Taking screenshot of Excel...")

    time.sleep(1)

    screenshot = pyautogui.screenshot()

    screenshot.save(
        screenshot_file
    )

    print()
    print("SCREENSHOT SAVED")
    print(screenshot_file)
    print()


# ============================================================
# 9. CLOSE EXCEL
# ============================================================

def close_excel():

    print("STEP 8: Closing Excel...")

    pyautogui.hotkey(
        "alt",
        "f4"
    )

    time.sleep(1.5)


# ============================================================
# 10. MAIN
# ============================================================

def main():

    print()
    print("==============================================")
    print("          DAILY REPORT AUTOMATION")
    print("==============================================")
    print()

    open_chrome()

    open_website()

    data = get_web_data()

    open_excel()

    enter_excel_data(data)

    take_screenshot()

    save_excel()

    close_excel()

    print()
    print("==============================================")
    print("             REPORT COMPLETED")
    print("==============================================")
    print()

    print("Excel:")
    print(excel_file)

    print()

    print("Screenshot:")
    print(screenshot_file)

    print()


# ============================================================
# RUN
# ============================================================

if __name__ == "__main__":
    main()
