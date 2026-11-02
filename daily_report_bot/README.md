# 📊 Project 1: Daily Report Bot (PyAutoGUI)

A **desktop automation** bot. It drives the keyboard like a person would: it opens Chrome on CoinDesk, copies the
page, pulls out the live Bitcoin price, highlights it with Ctrl+F as proof, takes a screenshot, closes Chrome and
writes a daily Excel report.

## 🎬 Demo

[▶ Watch the demo (27 s)](demo/daily_report_bot_demo.mp4): a real run recorded on screen, with the waiting time
sped up.

## How it works

```text
Win+R → chrome --profile-directory="Profile 1" --new-window --start-maximized <URL>
                            opens CoinDesk (fixed profile skips Chrome's profile picker)
wait for the page           10 s
Ctrl+A, Ctrl+C              copy the page text (pyperclip reads the clipboard)
regex                       "BTC #1 $77,291.77" → $77,291.77   (retries if the page isn't ready)
Ctrl+F <price>              Chrome scrolls to the price and highlights it
screenshot                  daily_report_<date>.png
Alt+F4                      close Chrome
openpyxl                    add a row to daily_report_<date>.xlsx
```

Example row:

| Date & Time | Bitcoin Price (USD) | Source | Comment |
|---|---|---|---|
| 2026-11-02 22:44:11 | $77,291.77 | CoinDesk | Market information captured successfully |

## Files

| File | Purpose |
|---|---|
| `daily_report_bot.py` | The bot |
| `find_coordinates.py` | Helper: prints the mouse position live, to find screen points |
| `screen_record.py` | Helper: records the screen to MP4 while the bot runs (Ctrl+C to stop) |
| `demo/daily_report_bot_demo.mp4` | Demo video |
| `requirements.txt` | pyautogui, pyperclip, pillow, openpyxl, opencv-python, numpy |

## Run

```bash
cd daily_report_bot
pip install -r requirements.txt
python daily_report_bot.py
```

* Needs Windows with **Google Chrome**. Set `CHROME_PROFILE` in the script to your profile folder
  (see `chrome://version` → Profile Path, e.g. `Default` or `Profile 1`).
* Don't touch the mouse or keyboard while it runs (about 20 seconds).
* Emergency stop: move the mouse to a screen corner (PyAutoGUI fail-safe).
* Reports are saved next to the script; running it again the same day adds a new row.

## What I learned

* Keyboard automation with PyAutoGUI, clipboard handling with pyperclip, data extraction with regex
* GUI automation is fragile: Chrome's profile picker or an app popup takes the keystrokes. Launching Chrome with a
  fixed profile, retrying the copy, and writing Excel with openpyxl (instead of clicking through Excel's dialogs)
  made the bot reliable
* Proving the result: Ctrl+F highlights the exact value that was captured before the screenshot
* PyAutoGUI's limits for web data, which is why Project 2 uses Playwright
