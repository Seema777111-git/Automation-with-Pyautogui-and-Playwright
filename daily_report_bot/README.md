# 📊 Project 1: Daily Report Bot (PyAutoGUI)

A **desktop automation** bot. It controls the mouse and keyboard like a person would: it opens Chrome, reads the
live Bitcoin price from CoinDesk, types it into a new Excel sheet, saves the report and takes a screenshot.

## How it works

```text
Win+R → "chrome"            open Chrome
Ctrl+T, Ctrl+L, paste URL   open coindesk.com/price/bitcoin
Ctrl+F "Bitcoin"            find the price row
Shift+Down/End, Ctrl+C      copy it, regex pulls out the $ price
Win+R → "excel"             open Excel, new workbook
type headers + data         Date & Time | Fetched Data | Comment
F12 → save                  daily_report_<date>.xlsx (next to the script)
screenshot                  daily_report_<date>.png
Alt+F4                      close Excel
```

## Files

| File | Purpose |
|---|---|
| `daily_report_bot.py` | The bot |
| `find_coordinates.py` | Helper: prints the mouse position live, to find screen points |
| `screen_record.py` | Helper: records the screen to MP4 while the bot runs (Ctrl+C to stop) |
| `requirements.txt` | pyautogui, pyperclip, pillow, opencv-python, numpy |

## Run

```bash
cd daily_report_bot
pip install -r requirements.txt
python daily_report_bot.py
```

* Needs Windows with **Chrome** and **Microsoft Excel** installed.
* Don't touch the mouse or keyboard while it runs (about 40 seconds).
* Emergency stop: move the mouse to a screen corner (PyAutoGUI fail-safe).
* Reports are saved in the same folder as the script.

## What I learned

* Keyboard and mouse automation with PyAutoGUI
* Clipboard handling with pyperclip and extracting data with regex
* Why timing matters: searching the page before it has loaded returns nothing,
  so the bot now waits for the page to load first
* PyAutoGUI's limits for web data, which is why Project 2 uses Playwright
