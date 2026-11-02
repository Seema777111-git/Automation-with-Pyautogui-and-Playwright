# ============================================================
# WHATSAPP PLAYWRIGHT AUTOMATION
# ============================================================
# REQUIREMENT:
#
# 1. Read contacts from Excel
# 2. Open WhatsApp Web
# 3. Login using saved persistent browser profile
# 4. For each contact:
#       - Search using phone number
#       - Open the contact chat
#       - Send personalized message
#       - Replace {name} with Excel name
#       - Wait until message is sent
#       - Take screenshot of sent message
#       - Extract last 3 messages received FROM contact
# 5. Move to next contact
# 6. Save final results to JSON and Excel
# ============================================================

import asyncio
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from playwright.async_api import async_playwright


# ============================================================
# PRINT OUTPUT IMMEDIATELY
# ============================================================

sys.stdout.reconfigure(line_buffering=True)


# ============================================================
# SECTION 1 - PROJECT SETTINGS
# ============================================================

# Folder that contains this script (works on any PC)
BASE_DIR = Path(__file__).resolve().parent

CONTACT_FILE = BASE_DIR / "Contact.xlsx"

WHATSAPP_URL = "https://web.whatsapp.com"

PROFILE_FOLDER = BASE_DIR / "whatsapp_profile"

SCREENSHOT_FOLDER = BASE_DIR / "screenshots"

DEFAULT_MESSAGE = (
    "Hello {name}, this is a test message."
)


# ============================================================
# CREATE REQUIRED FOLDERS
# ============================================================

BASE_DIR.mkdir(
    parents=True,
    exist_ok=True
)

PROFILE_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)

SCREENSHOT_FOLDER.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# REPORT FILES
# ============================================================

today = datetime.now().strftime("%Y-%m-%d")

REPORT_JSON = (
    BASE_DIR / f"whatsapp_report_{today}.json"
)

REPORT_EXCEL = (
    BASE_DIR / f"whatsapp_report_{today}.xlsx"
)


# ============================================================
# WHATSAPP SEARCH SELECTORS
# ============================================================

SEARCH_BOX_SELECTORS = [

    '#side div[contenteditable="true"][data-tab="3"]',

    'div[contenteditable="true"]'
    '[aria-label="Search input textbox"]',

    '#side div[contenteditable="true"]',

    'div[contenteditable="true"]'
    '[role="textbox"]',

    'input[placeholder*="Search"]',

]


# ============================================================
# MESSAGE COMPOSE BOX SELECTORS
# ============================================================

MESSAGE_BOX_SELECTORS = [

    'div[contenteditable="true"]'
    '[aria-placeholder="Type a message"]',

    'div[contenteditable="true"]'
    '[data-placeholder="Type a message"]',

    'footer div[contenteditable="true"][data-tab]',

    '#main footer div[contenteditable="true"]',

    'div[aria-label="Type a message"]',

    'footer div[contenteditable="true"]',

]


# ============================================================
# SEND BUTTON SELECTORS
# ============================================================

SEND_BUTTON_SELECTORS = [

    'button[aria-label="Send"]',

    'button[data-testid="send"]',

    'span[data-icon="send"]',

]


# ============================================================
# SEARCH RESULT SELECTORS
# ============================================================

SEARCH_RESULT_SELECTORS = [

    'div[data-testid="cell-frame-container"]',

    '#pane-side div[role="listitem"]',

    'div[aria-label="Search results."]'
    ' div[role="listitem"]',

    'div[role="grid"] div[role="row"]',

]


# ============================================================
# INCOMING MESSAGE SELECTORS
# ============================================================

INCOMING_MESSAGE_SELECTORS = [

    'div.message-in',

    'div[class*="message-in"]',

]


# ============================================================
# SENT MESSAGE SELECTORS
# ============================================================

SENT_MESSAGE_SELECTORS = [

    'div.message-out',

    'div[class*="message-out"]',

]


# ============================================================
# SENT TICK SELECTORS
# ============================================================

SENT_TICK_SELECTORS = [

    'span[data-icon="msg-check"]',

    'span[data-icon="msg-dblcheck"]',

    'span[data-icon="msg-dblcheck-ack"]',

    'span[aria-label="Sent"]',

    'span[aria-label="Delivered"]',

    'span[aria-label="Read"]',

]


# ============================================================
# DEBUG COUNTER
# ============================================================

DEBUG_COUNTER = 0


# ============================================================
# HELPER - DEBUG SNAPSHOT
# ============================================================

async def save_debug_snapshot(
    page,
    reason
):

    global DEBUG_COUNTER

    DEBUG_COUNTER += 1

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    safe_reason = re.sub(
        r"[^A-Za-z0-9_-]",
        "_",
        reason
    )

    filename = (
        f"debug_{DEBUG_COUNTER:03d}_"
        f"{safe_reason}_"
        f"{timestamp}"
    )

    screenshot_path = (
        SCREENSHOT_FOLDER /
        f"{filename}.png"
    )

    html_path = (
        SCREENSHOT_FOLDER /
        f"{filename}.html"
    )

    try:

        await page.screenshot(
            path=str(screenshot_path),
            full_page=True
        )

        print(
            "Debug screenshot saved:",
            screenshot_path
        )

    except Exception as error:

        print(
            "Could not save debug screenshot:",
            error
        )

    try:

        html = await page.content()

        html_path.write_text(
            html,
            encoding="utf-8"
        )

        print(
            "Debug HTML saved:",
            html_path
        )

    except Exception as error:

        print(
            "Could not save debug HTML:",
            error
        )


# ============================================================
# HELPER - RESET WHATSAPP STATE
# ============================================================

async def reset_to_chat_list(page):

    try:

        await page.keyboard.press("Escape")

        await page.wait_for_timeout(500)

        await page.keyboard.press("Escape")

        await page.wait_for_timeout(500)

    except Exception:

        pass


# ============================================================
# HELPER - NORMALIZE PHONE
# ============================================================

def normalize_phone(phone):

    digits = re.sub(
        r"\D",
        "",
        str(phone)
    )

    # Indian 10 digit number
    if len(digits) == 10:

        digits = "91" + digits

    return "+" + digits


# ============================================================
# SECTION 2 - READ CONTACTS FROM EXCEL
# ============================================================

def read_contacts():

    print()
    print("==========================================")
    print("SECTION 2 - READ CONTACTS FROM EXCEL")
    print("==========================================")

    # --------------------------------------------------------
    # CHECK FILE
    # --------------------------------------------------------

    if not CONTACT_FILE.exists():

        raise FileNotFoundError(
            f"\nExcel file not found:\n{CONTACT_FILE}"
        )

    # --------------------------------------------------------
    # READ EXCEL
    # --------------------------------------------------------

    df = pd.read_excel(
        CONTACT_FILE
    )

    print(
        "Excel columns:",
        list(df.columns)
    )

    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    required_columns = [
        "Name",
        "Phone",
        "Message"
    ]

    for column in required_columns:

        if column not in df.columns:

            raise ValueError(
                f"Missing Excel column: {column}"
            )

    # --------------------------------------------------------
    # BUILD CONTACT LIST
    # --------------------------------------------------------

    contacts = []

    for _, row in df.iterrows():

        name = str(
            row["Name"]
        ).strip()

        phone_raw = str(
            row["Phone"]
        ).strip()

        message = row["Message"]

        # ----------------------------------------------------
        # SKIP INVALID ROW
        # ----------------------------------------------------

        if not name:

            continue

        if not phone_raw:

            continue

        if phone_raw.lower() == "nan":

            continue

        # ----------------------------------------------------
        # NORMALIZE PHONE
        # ----------------------------------------------------

        phone = normalize_phone(
            phone_raw
        )

        # ----------------------------------------------------
        # MESSAGE
        # ----------------------------------------------------

        if pd.isna(message):

            message = DEFAULT_MESSAGE

        message = str(
            message
        ).strip()

        if not message:

            message = DEFAULT_MESSAGE

        if message.lower() == "nan":

            message = DEFAULT_MESSAGE

        # ----------------------------------------------------
        # SAVE CONTACT
        # ----------------------------------------------------

        contacts.append(
            {
                "name": name,
                "phone": phone,
                "message": message
            }
        )

    # --------------------------------------------------------
    # DISPLAY CONTACTS
    # --------------------------------------------------------

    print()
    print(
        "Contacts found:",
        len(contacts)
    )

    print()

    for contact in contacts:

        print(
            f'{contact["name"]} | '
            f'{contact["phone"]} | '
            f'{contact["message"]}'
        )

    print()

    return contacts


# ============================================================
# SECTION 3 - OPEN WHATSAPP WEB
# ============================================================

async def open_whatsapp(page):

    print()
    print("==========================================")
    print("SECTION 3 - OPEN WHATSAPP WEB")
    print("==========================================")

    print()
    print("Opening WhatsApp Web...")

    await page.goto(
        WHATSAPP_URL,
        wait_until="domcontentloaded"
    )

    print(
        "WhatsApp title:",
        await page.title()
    )

    # Give WhatsApp time to load
    await page.wait_for_timeout(
        5000
    )

    print()
    print(
        "Checking WhatsApp login..."
    )

    # --------------------------------------------------------
    # LOGIN INDICATORS
    # --------------------------------------------------------

    login_indicators = [
        '#pane-side',
        'div[aria-label="Chat list"]',
        '#side'
    ]

    for selector in login_indicators:

        try:

            await page.locator(
                selector
            ).first.wait_for(
                state="visible",
                timeout=5000
            )

            print()
            print("==========================================")
            print("WHATSAPP ALREADY LOGGED IN")
            print("==========================================")
            print()

            return True

        except Exception:

            continue

    # --------------------------------------------------------
    # QR LOGIN
    # --------------------------------------------------------

    print()
    print("==========================================")
    print("QR CODE LOGIN REQUIRED")
    print("==========================================")

    print()
    print(
        "Scan the QR code using your phone."
    )

    print(
        "WhatsApp -> Linked Devices -> "
        "Link a Device"
    )

    print()
    print(
        "Waiting for login..."
    )

    # --------------------------------------------------------
    # WAIT FOR LOGIN
    # --------------------------------------------------------

    try:

        await page.locator(
            '#pane-side, '
            'div[aria-label="Chat list"], '
            '#side'
        ).first.wait_for(
            state="visible",
            timeout=120000
        )

        print()
        print("==========================================")
        print("WHATSAPP LOGIN SUCCESSFUL")
        print("==========================================")
        print()

        return True

    except Exception as error:

        print()
        print("==========================================")
        print("WHATSAPP LOGIN FAILED")
        print("==========================================")

        print(
            "Error:",
            repr(error)
        )

        await save_debug_snapshot(
            page,
            "login_failed"
        )

        return False


# ============================================================
# HELPER - FIND SEARCH BOX
# ============================================================

async def find_search_box(page):

    for selector in SEARCH_BOX_SELECTORS:

        locator = page.locator(
            selector
        ).first

        try:

            await locator.wait_for(
                state="visible",
                timeout=3000
            )

            print(
                "[search box matched]",
                selector
            )

            return locator

        except Exception:

            continue

    return None


# ============================================================
# HELPER - FIND SEARCH RESULTS
# ============================================================

async def find_search_results(page):

    for selector in SEARCH_RESULT_SELECTORS:

        locator = page.locator(
            selector
        )

        try:

            await locator.first.wait_for(
                state="visible",
                timeout=3000
            )

            count = await locator.count()

            if count > 0:

                print(
                    "[search results matched]",
                    selector,
                    "->",
                    count
                )

                return locator

        except Exception:

            continue

    return None


# ============================================================
# HELPERS - MAKE SURE WE MESSAGE THE RIGHT PERSON
# ============================================================

def result_matches_contact(
    text,
    excel_name,
    target_last10
):
    """True if the text shows the contact's name or phone number."""

    text = str(text or "")

    if excel_name and excel_name.lower() in text.lower():
        return True

    text_digits = re.sub(r"\D", "", text)

    return bool(target_last10) and target_last10 in text_digits


async def get_chat_header_text(page):
    """Title of the chat that is open now (name or number)."""

    for selector in [
        "#main header span[dir='auto']",
        "#main header",
    ]:

        try:

            locator = page.locator(selector).first

            if await locator.count() > 0:

                text = (await locator.inner_text()).strip()

                if text:
                    return text

        except Exception:
            continue

    return ""


# ============================================================
# STEP 4 - SEARCH CONTACT BY PHONE
# ============================================================

async def search_contact(
    page,
    name,
    phone
):

    print()
    print("==========================================")
    print("STEP 4 - SEARCH CONTACT")
    print("==========================================")

    excel_name = str(
        name
    ).strip()

    excel_phone = str(
        phone
    ).strip()

    print(
        "Excel Name :",
        excel_name
    )

    print(
        "Excel Phone:",
        excel_phone
    )

    # --------------------------------------------------------
    # PHONE NORMALIZATION
    # --------------------------------------------------------

    digits = re.sub(
        r"\D",
        "",
        excel_phone
    )

    target_last10 = digits[-10:]

    search_number = "+91" + target_last10

    print(
        "Search Number:",
        search_number
    )

    # --------------------------------------------------------
    # RESET
    # --------------------------------------------------------

    await reset_to_chat_list(
        page
    )

    # --------------------------------------------------------
    # FIND SEARCH BOX
    # --------------------------------------------------------

    search_box = await find_search_box(
        page
    )

    if search_box is None:

        print()
        print(
            "Search box not found."
        )

        await save_debug_snapshot(
            page,
            f"search_box_missing_{excel_name}"
        )

        return False

    try:

        # ----------------------------------------------------
        # CLEAR SEARCH
        # ----------------------------------------------------

        await search_box.click()

        await page.keyboard.press(
            "Control+A"
        )

        await page.keyboard.press(
            "Backspace"
        )

        # ----------------------------------------------------
        # SEARCH PHONE
        # ----------------------------------------------------

        print()
        print(
            "Searching phone:",
            search_number
        )

        await search_box.fill(
            search_number
        )

        # ----------------------------------------------------
        # WAIT FOR RESULTS
        # ----------------------------------------------------

        await page.wait_for_timeout(
            2500
        )

        results = await find_search_results(
            page
        )

        if results is None:

            print()
            print("==========================================")
            print("CONTACT NOT FOUND")
            print("==========================================")

            print(
                "Name :",
                excel_name
            )

            print(
                "Phone:",
                excel_phone
            )

            await save_debug_snapshot(
                page,
                f"contact_not_found_{excel_name}"
            )

            return False

        count = await results.count()

        print(
            "Search result count:",
            count
        )

        # ----------------------------------------------------
        # CHECK RESULTS
        # ----------------------------------------------------

        for index in range(count):

            result = results.nth(
                index
            )

            try:

                if not await result.is_visible():

                    continue

                result_text = await result.inner_text()

                print()
                print(
                    f"Result {index + 1}"
                )

                print(
                    "------------------------------------------"
                )

                print(
                    result_text
                )

                # ------------------------------------------------
                # IGNORE GROUPS
                # ------------------------------------------------

                lower_text = (
                    result_text.lower()
                )

                if (
                    "group" in lower_text
                    or
                    "participants" in lower_text
                    or
                    "members" in lower_text
                ):

                    print(
                        "Group detected - skipping."
                    )

                    continue

                # ------------------------------------------------
                # SAFETY: ONLY OPEN A RESULT THAT MATCHES
                # (name or last 10 digits of the phone)
                # ------------------------------------------------

                if not result_matches_contact(
                    result_text,
                    excel_name,
                    target_last10
                ):

                    print(
                        "Not this contact - skipping."
                    )

                    continue

                # ------------------------------------------------
                # OPEN RESULT
                # ------------------------------------------------

                print(
                    "Opening contact..."
                )

                await result.click()

                # ------------------------------------------------
                # WAIT FOR CHAT
                # ------------------------------------------------

                await page.locator(
                    "#main"
                ).wait_for(
                    state="visible",
                    timeout=10000
                )

                # ------------------------------------------------
                # SAFETY: CONFIRM THE OPEN CHAT IS THE RIGHT ONE
                # ------------------------------------------------

                header_text = await get_chat_header_text(
                    page
                )

                print(
                    "Open chat header:",
                    header_text
                )

                if not result_matches_contact(
                    header_text,
                    excel_name,
                    target_last10
                ):

                    print(
                        "Open chat does not match this contact - "
                        "NOT sending."
                    )

                    await save_debug_snapshot(
                        page,
                        f"wrong_chat_{excel_name}"
                    )

                    return False

                print()
                print("==========================================")
                print("CONTACT FOUND")
                print("==========================================")

                print(
                    "Name :",
                    excel_name
                )

                print(
                    "Phone:",
                    excel_phone
                )

                print(
                    "=========================================="
                )

                return True

            except Exception as error:

                print(
                    f"Error opening result {index + 1}:",
                    error
                )

                continue

        # ----------------------------------------------------
        # NO RESULT OPENED
        # ----------------------------------------------------

        print()
        print("==========================================")
        print("CONTACT NOT FOUND")
        print("==========================================")

        print(
            "Name :",
            excel_name
        )

        print(
            "Phone:",
            excel_phone
        )

        return False

    except Exception as error:

        print()
        print("==========================================")
        print("STEP 4 ERROR")
        print("==========================================")

        print(
            "Name :",
            excel_name
        )

        print(
            "Phone:",
            excel_phone
        )

        print(
            "Error:",
            error
        )

        await save_debug_snapshot(
            page,
            f"step4_error_{excel_name}"
        )

        return False


# ============================================================
# HELPER - FIND MESSAGE BOX
# ============================================================

async def find_message_box(page):

    for selector in MESSAGE_BOX_SELECTORS:

        locator = page.locator(
            selector
        ).last

        try:

            await locator.wait_for(
                state="visible",
                timeout=3000
            )

            print(
                "[message box matched]",
                selector
            )

            return locator

        except Exception:

            continue

    print(
        "Message compose box not found."
    )

    await save_debug_snapshot(
        page,
        "message_box_missing"
    )

    return None


# ============================================================
# STEP 5 - SEND PERSONALIZED MESSAGE
# ============================================================

async def send_message(
    page,
    name,
    phone,
    message_template
):

    print()
    print("==========================================")
    print("STEP 5 - SEND MESSAGE")
    print("==========================================")

    excel_name = str(
        name
    ).strip()

    excel_phone = str(
        phone
    ).strip()

    # --------------------------------------------------------
    # PERSONALIZE MESSAGE
    # --------------------------------------------------------

    personalized_message = (
        str(message_template)
        .replace(
            "{name}",
            excel_name
        )
    )

    print(
        "To      :",
        excel_name
    )

    print(
        "Phone   :",
        excel_phone
    )

    print(
        "Message :",
        personalized_message
    )

    try:

        # ----------------------------------------------------
        # FIND MESSAGE BOX
        # ----------------------------------------------------

        message_box = await find_message_box(
            page
        )

        if message_box is None:

            return False

        # ----------------------------------------------------
        # CLEAR MESSAGE BOX
        # ----------------------------------------------------

        await message_box.click()

        await page.keyboard.press(
            "Control+A"
        )

        await page.keyboard.press(
            "Backspace"
        )

        # ----------------------------------------------------
        # TYPE MESSAGE
        # ----------------------------------------------------

        lines = personalized_message.split(
            "\n"
        )

        for index, line in enumerate(lines):

            await message_box.type(
                line,
                delay=15
            )

            if index < len(lines) - 1:

                await page.keyboard.press(
                    "Shift+Enter"
                )

        # ----------------------------------------------------
        # SEND MESSAGE
        # ----------------------------------------------------

        sent_using_button = False

        for selector in SEND_BUTTON_SELECTORS:

            try:

                send_button = page.locator(
                    selector
                ).last

                await send_button.wait_for(
                    state="visible",
                    timeout=2000
                )

                await send_button.click()

                sent_using_button = True

                print(
                    "Message sent using Send button."
                )

                break

            except Exception:

                continue

        # ----------------------------------------------------
        # FALLBACK ENTER
        # ----------------------------------------------------

        if not sent_using_button:

            print(
                "Send button not found."
            )

            print(
                "Using Enter key..."
            )

            await page.keyboard.press(
                "Enter"
            )

        # ----------------------------------------------------
        # WAIT FOR SENT TICK
        # ----------------------------------------------------

        sent_confirmed = False

        print(
            "Waiting for message to be sent..."
        )

        for selector in SENT_TICK_SELECTORS:

            try:

                await page.locator(
                    selector
                ).last.wait_for(
                    state="visible",
                    timeout=5000
                )

                print(
                    "[sent confirmation]",
                    selector
                )

                sent_confirmed = True

                break

            except Exception:

                continue

        # ----------------------------------------------------
        # SMALL WAIT
        # ----------------------------------------------------

        await page.wait_for_timeout(
            1000
        )

        # ----------------------------------------------------
        # SCREENSHOT
        # ----------------------------------------------------

        safe_name = re.sub(
            r"[^A-Za-z0-9_-]",
            "_",
            excel_name
        )

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        screenshot_path = (
            SCREENSHOT_FOLDER /
            f"{safe_name}_{timestamp}.png"
        )

        screenshot_saved = False

        # ----------------------------------------------------
        # TRY SENT MESSAGE BUBBLE
        # ----------------------------------------------------

        for selector in SENT_MESSAGE_SELECTORS:

            try:

                sent_messages = page.locator(
                    selector
                )

                count = await sent_messages.count()

                if count == 0:

                    continue

                last_message = sent_messages.last

                await last_message.wait_for(
                    state="visible",
                    timeout=3000
                )

                await last_message.screenshot(
                    path=str(screenshot_path)
                )

                screenshot_saved = True

                print(
                    "Screenshot of sent message saved:",
                    screenshot_path
                )

                break

            except Exception:

                continue

        # ----------------------------------------------------
        # FALLBACK CHAT SCREENSHOT
        # ----------------------------------------------------

        if not screenshot_saved:

            try:

                await page.locator(
                    "#main"
                ).screenshot(
                    path=str(screenshot_path)
                )

                screenshot_saved = True

                print(
                    "Chat screenshot saved:",
                    screenshot_path
                )

            except Exception:

                await page.screenshot(
                    path=str(screenshot_path),
                    full_page=True
                )

                screenshot_saved = True

                print(
                    "Full page screenshot saved:",
                    screenshot_path
                )

        # ----------------------------------------------------
        # RESULT
        # ----------------------------------------------------

        print()
        print("==========================================")
        print("MESSAGE SENT")
        print("==========================================")

        print(
            "Name      :",
            excel_name
        )

        print(
            "Phone     :",
            excel_phone
        )

        print(
            "Confirmed  :",
            sent_confirmed
        )

        print(
            "Screenshot:",
            screenshot_path
        )

        print(
            "=========================================="
        )

        return True

    except Exception as error:

        print()
        print("==========================================")
        print("MESSAGE SEND FAILED")
        print("==========================================")

        print(
            "Name :",
            excel_name
        )

        print(
            "Phone:",
            excel_phone
        )

        print(
            "Error:",
            error
        )

        await save_debug_snapshot(
            page,
            f"send_error_{excel_name}"
        )

        return False


# ============================================================
# STEP 6 - EXTRACT LAST 3 INCOMING MESSAGES
# ============================================================

async def extract_last_messages(
    page,
    name,
    phone,
    count=3
):

    print()
    print("==========================================")
    print("STEP 6 - EXTRACT LAST 3 MESSAGES")
    print("==========================================")

    excel_name = str(
        name
    ).strip()

    excel_phone = str(
        phone
    ).strip()

    print(
        "Contact:",
        excel_name
    )

    print(
        "Phone  :",
        excel_phone
    )

    try:

        incoming_messages = None

        # ----------------------------------------------------
        # FIND INCOMING MESSAGE BUBBLES
        # ----------------------------------------------------

        for selector in INCOMING_MESSAGE_SELECTORS:

            locator = page.locator(
                selector
            )

            try:

                await locator.first.wait_for(
                    state="visible",
                    timeout=3000
                )

                if await locator.count() > 0:

                    incoming_messages = locator

                    print(
                        "[incoming message selector matched]",
                        selector
                    )

                    break

            except Exception:

                continue

        # ----------------------------------------------------
        # NO MESSAGES
        # ----------------------------------------------------

        if incoming_messages is None:

            print(
                "No incoming messages found."
            )

            return []

        # ----------------------------------------------------
        # TOTAL MESSAGES
        # ----------------------------------------------------

        total = await incoming_messages.count()

        print(
            "Total incoming messages:",
            total
        )

        # ----------------------------------------------------
        # GET LAST 3
        # ----------------------------------------------------

        start_index = max(
            0,
            total - count
        )

        messages = []

        for index in range(
            start_index,
            total
        ):

            try:

                bubble = incoming_messages.nth(
                    index
                )

                text = await bubble.inner_text()

                text = text.strip()

                if text:

                    messages.append(
                        text
                    )

            except Exception as error:

                print(
                    f"Error reading message {index}:",
                    error
                )

        # ----------------------------------------------------
        # DISPLAY
        # ----------------------------------------------------

        print()
        print(
            f"Last {len(messages)} "
            f"message(s) from {excel_name}:"
        )

        print(
            "------------------------------------------"
        )

        for index, message in enumerate(
            messages,
            start=1
        ):

            print(
                f"{index}. {message}"
            )

        print(
            "=========================================="
        )

        return messages

    except Exception as error:

        print()
        print("==========================================")
        print("MESSAGE EXTRACTION FAILED")
        print("==========================================")

        print(
            "Name :",
            excel_name
        )

        print(
            "Phone:",
            excel_phone
        )

        print(
            "Error:",
            error
        )

        return []


# ============================================================
# SAVE JSON REPORT
# ============================================================

def save_json_report(summary):

    try:

        with open(
            REPORT_JSON,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                summary,
                file,
                indent=2,
                ensure_ascii=False
            )

        print()
        print(
            "JSON report saved:",
            REPORT_JSON
        )

    except Exception as error:

        print(
            "Could not save JSON report:",
            error
        )


# ============================================================
# SAVE EXCEL REPORT
# ============================================================

def save_excel_report(summary):

    try:

        rows = []

        for item in summary:

            rows.append(
                {
                    "Name": item["name"],
                    "Phone": item["phone"],
                    "Contact Found": item["found"],
                    "Message Sent": item["sent"],
                    "Messages Extracted": len(
                        item["last_messages"]
                    ),
                    "Last Message 1": (
                        item["last_messages"][0]
                        if len(item["last_messages"]) > 0
                        else ""
                    ),
                    "Last Message 2": (
                        item["last_messages"][1]
                        if len(item["last_messages"]) > 1
                        else ""
                    ),
                    "Last Message 3": (
                        item["last_messages"][2]
                        if len(item["last_messages"]) > 2
                        else ""
                    ),
                    "Screenshot": item.get(
                        "screenshot",
                        ""
                    )
                }
            )

        df = pd.DataFrame(
            rows
        )

        df.to_excel(
            REPORT_EXCEL,
            index=False
        )

        print(
            "Excel report saved:",
            REPORT_EXCEL
        )

    except Exception as error:

        print(
            "Could not save Excel report:",
            error
        )


# ============================================================
# MAIN
# ============================================================

async def main():

    print()
    print("==========================================")
    print("WHATSAPP AUTOMATION BOT")
    print("==========================================")

    print()
    print(
        "Project folder:",
        BASE_DIR
    )

    print(
        "Contacts file:",
        CONTACT_FILE
    )

    print()

    # ========================================================
    # READ CONTACTS
    # ========================================================

    try:

        contacts = read_contacts()

    except Exception as error:

        print()
        print(
            "Could not read contacts:"
        )

        print(
            error
        )

        return

    if not contacts:

        print(
            "No valid contacts found."
        )

        return

    # ========================================================
    # START PLAYWRIGHT
    # ========================================================

    async with async_playwright() as playwright:

        print()
        print(
            "Starting Chromium..."
        )

        context = await playwright.chromium.launch_persistent_context(
            user_data_dir=str(
                PROFILE_FOLDER
            ),
            headless=False
        )

        # ----------------------------------------------------
        # GET PAGE
        # ----------------------------------------------------

        if context.pages:

            page = context.pages[0]

        else:

            page = await context.new_page()

        # ====================================================
        # SECTION 3 - OPEN WHATSAPP
        # ====================================================

        whatsapp_ready = await open_whatsapp(
            page
        )

        if not whatsapp_ready:

            print(
                "WhatsApp is not ready."
            )

            await context.close()

            return

        print()
        print("Section 3 completed successfully.")

        # ====================================================
        # CONTACT PROCESSING
        # ====================================================

        summary = []

        for contact_number, contact in enumerate(
            contacts,
            start=1
        ):

            print()
            print()
            print("==========================================")
            print(
                f"PROCESSING CONTACT "
                f"{contact_number}/{len(contacts)}"
            )
            print("==========================================")

            name = contact["name"]

            phone = contact["phone"]

            message = contact["message"]

            found = False

            sent = False

            last_messages = []

            screenshot = ""

            # =================================================
            # STEP 4
            # =================================================

            found = await search_contact(
                page,
                name,
                phone
            )

            # =================================================
            # STEP 5
            # =================================================

            if found:

                sent = await send_message(
                    page,
                    name,
                    phone,
                    message
                )

                # ------------------------------------------------
                # FIND LATEST SCREENSHOT
                # ------------------------------------------------

                if sent:

                    safe_name = re.sub(
                        r"[^A-Za-z0-9_-]",
                        "_",
                        str(name)
                    )

                    matching_files = sorted(
                        SCREENSHOT_FOLDER.glob(
                            f"{safe_name}_*.png"
                        )
                    )

                    if matching_files:

                        screenshot = str(
                            matching_files[-1]
                        )

                # =================================================
                # STEP 6
                # =================================================

                last_messages = (
                    await extract_last_messages(
                        page,
                        name,
                        phone,
                        count=3
                    )
                )

            else:

                print()
                print(
                    "Skipping message and extraction."
                )

                print(
                    "Contact was not found:",
                    name
                )

            # =================================================
            # SAVE CONTACT RESULT
            # =================================================

            summary.append(
                {
                    "name": name,
                    "phone": phone,
                    "found": found,
                    "sent": sent,
                    "screenshot": screenshot,
                    "last_messages": last_messages
                }
            )

            # =================================================
            # NEXT CONTACT
            # =================================================

            print()
            print(
                "Preparing for next contact..."
            )

            await reset_to_chat_list(
                page
            )

            await page.wait_for_timeout(
                1500
            )

        # ====================================================
        # FINAL SUMMARY
        # ====================================================

        print()
        print()
        print("==========================================")
        print("FINAL RUN SUMMARY")
        print("==========================================")

        for item in summary:

            if item["sent"]:

                status = "SENT"

            elif item["found"]:

                status = "FOUND - NOT SENT"

            else:

                status = "CONTACT NOT FOUND"

            print(
                f'{item["name"]:15} | '
                f'{item["phone"]:15} | '
                f'{status:20} | '
                f'Last messages: '
                f'{len(item["last_messages"])}'
            )

        print(
            "=========================================="
        )

        # ====================================================
        # SAVE REPORTS
        # ====================================================

        save_json_report(
            summary
        )

        save_excel_report(
            summary
        )

        # ====================================================
        # CLOSE BROWSER
        # ====================================================

        print()
        print(
            "Closing WhatsApp browser..."
        )

        await context.close()

        print()
        print("==========================================")
        print("AUTOMATION COMPLETED")
        print("==========================================")


# ============================================================
# PROGRAM START
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )