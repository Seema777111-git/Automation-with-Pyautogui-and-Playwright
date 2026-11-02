# 🤖 Automation with PyAutoGUI and Playwright

Two Python automation projects: one controls the **desktop** (mouse and keyboard), the other controls a **web
browser** (page elements).

| # | Project | Tool | What it automates |
|---|---|---|---|
| 1 | [**Daily Report Bot**](daily_report_bot) | PyAutoGUI | Opens Chrome on CoinDesk, copies the live Bitcoin price, highlights it as proof, saves a screenshot and a daily Excel report |
| 2 | [**WhatsApp Message Bot**](Play_wright_bot) | Playwright | Reads contacts from Excel, sends each a personalised WhatsApp Web message, screenshots it, reads their last 3 replies and saves a report |

🎬 **Demos (real runs):**
[Daily Report Bot (27 s)](daily_report_bot/demo/daily_report_bot_demo.mp4) ·
[WhatsApp Message Bot (29 s)](Play_wright_bot/demo/whatsapp_bot_demo.mp4)

```text
Automation-with-Pyautogui-and-Playwright/
├── daily_report_bot/          Project 1 · PyAutoGUI
│   ├── daily_report_bot.py
│   ├── find_coordinates.py
│   ├── screen_record.py
│   ├── demo/daily_report_bot_demo.mp4
│   ├── requirements.txt
│   └── README.md
│
└── Play_wright_bot/           Project 2 · Playwright
    ├── playwright_assign.py
    ├── Contact.sample.xlsx
    ├── demo/whatsapp_bot_demo.mp4
    ├── requirements.txt
    └── README.md
```

Each folder has its own README with the steps, files and how to run it.

---

## PyAutoGUI vs Playwright

| | PyAutoGUI (Project 1) | Playwright (Project 2) |
|---|---|---|
| Controls | Mouse, keyboard, whole screen | The browser and its page elements |
| Finds things by | Keyboard shortcuts, screen position | CSS selectors, text, roles |
| Works with | Any desktop app (Chrome, Excel…) | Web pages |
| Reliability | Depends on timing and screen layout | Waits for elements; much more reliable |
| Runs in background | No, it takes over the PC | Yes (can run headless) |
| Best for | Desktop apps with no API | Web apps and browser testing |

```text
PyAutoGUI:   Python → mouse + keyboard → screen → any app
Playwright:  Python → browser → page elements (selectors) → web app
```

---

## Quick start

```bash
git clone https://github.com/Seema777111-git/Automation-with-Pyautogui-and-Playwright.git
cd Automation-with-Pyautogui-and-Playwright
python -m venv .venv
.venv\Scripts\activate
```

**Project 1**
```bash
cd daily_report_bot
pip install -r requirements.txt
python daily_report_bot.py
```

**Project 2**
```bash
cd Play_wright_bot
pip install -r requirements.txt
playwright install chromium
copy Contact.sample.xlsx Contact.xlsx      # add your contacts
python playwright_assign.py
```

---

## Skills shown

* Desktop automation: keyboard shortcuts, clipboard, screenshots (PyAutoGUI, pyperclip)
* Browser automation: persistent login, selectors, waits, debugging snapshots (Playwright)
* Excel input and output (pandas, openpyxl)
* Data extraction with regex
* Safe automation: checking the target before an action that can't be undone

## 👩‍💻 Author

**Seema** · Python Automation | AI | Generative AI | RAG | AI Agents ·
[GitHub](https://github.com/Seema777111-git)
