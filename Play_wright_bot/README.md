# 💬 Project 2: WhatsApp Message Bot (Playwright)

A **browser automation** bot. It reads a contact list from Excel, opens WhatsApp Web with Playwright, sends each
person a personalised message, takes a screenshot as proof, reads their last 3 replies, and saves a report.

## How it works

```text
Contact.xlsx (Name, Phone, Message)
        ↓
Read and normalise phone numbers (+91 XXXXXXXXXX)
        ↓
Open WhatsApp Web in Chromium (saved login profile; QR scan only the first time)
        ↓
For each contact:
   search by phone number
   → open only a result that shows the contact's name or number
   → confirm the open chat's header matches before sending
   → type the message ({name} is replaced), press Send
   → screenshot of the sent message
   → read the last 3 incoming messages
        ↓
whatsapp_report_<date>.json + .xlsx
```

## Safety check

The bot never sends a message unless the chat it opened really belongs to the contact. Search results that don't
show the contact's name or number are skipped, and the open chat's header is checked again before typing. If it
doesn't match, the contact is reported as **not found** and nothing is sent.

## Files

| File | Purpose |
|---|---|
| `playwright_assign.py` | The bot |
| `Contact.sample.xlsx` | Example contact list (fake numbers). Copy it to `Contact.xlsx` and add your contacts |
| `requirements.txt` | playwright, pandas, openpyxl |

## Run

```bash
cd Play_wright_bot
pip install -r requirements.txt
playwright install chromium
copy Contact.sample.xlsx Contact.xlsx      # then edit Contact.xlsx
python playwright_assign.py
```

* First run: scan the QR code with WhatsApp on your phone (Linked devices → Link a device). Wait until your chats
  have loaded.
* The login is saved in `whatsapp_profile/`. **Never upload that folder**; it is your WhatsApp session.
  `Contact.xlsx`, `screenshots/` and the reports are also kept out of Git by `.gitignore`.
* Leave the `Message` column empty to use the default message.

## What I learned

* Playwright: persistent browser profiles, selectors with fallbacks, waiting for elements
* Reading and writing Excel with pandas
* Debugging with automatic screenshots and HTML snapshots
* Verifying the target before taking an action that can't be undone
