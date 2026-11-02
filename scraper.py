
import base64
import html
import io
import json
import os
import re
import sys
import threading
import time
from datetime import datetime, timedelta, timezone
from urllib.parse import quote_plus, urljoin

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from playwright.sync_api import sync_playwright

# ---------------- CONFIGURATION ----------------

# Every .xlsx in INPUT_DIR is a query file, processed in file-name order
# (name them 01_..., 02_... to control the order). Each file is finished
# before the next one is started.
INPUT_DIR = "input"
OUTPUT_DIR = "output"
API_CONFIG_FILE = "api_config.json"
SEARCH_URL = "https://www.google.com/maps/search/{}/"
XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

MAX_LISTINGS = 200
SCROLL_PAUSE = 2
MAX_IDLE_SCROLLS = 4

QUERY_COLUMN = "Google Maps Search Query"
POLL_INTERVAL = 600  # seconds between checks of the input files (10 min)
UPLOAD_RETRY_DELAY = 600  # seconds before re-hitting a failed upload
SCRAPE_RETRY_DELAY = 600  # seconds before re-scraping a failed query
FILE_OPEN_RECHECK = 30  # seconds between checks of an input file open in Excel

# The lead API is hit only between these hours (IST, 24h clock); scraping
# runs all day and completed rows wait for the window. 22 -> 6 means
# 10:00 PM to 6:00 AM. IST has no daylight saving, so a fixed offset is
# exact whatever the machine's time zone is.
IST = timezone(timedelta(hours=5, minutes=30), "IST")
UPLOAD_START_HOUR = 22
UPLOAD_END_HOUR = 6

# Selectors matched to the HTML you shared
CARD_SELECTOR = 'div[role="article"].Nv2PK'
NAME_SELECTOR = 'a.hfpxzc[aria-label]'
PHONE_SELECTOR = '.UsdlK'
ADDRESS_SELECTOR = '.W4Efsd'
WEBSITE_SELECTOR = 'a[data-value="Website"][href]'
RATING_SELECTOR = 'span.MW4etd'

# Copied from the input file row into every output record.
QUERY_INFO_COLUMNS = [
    "Country",
    "State",
    "City / Area",
    "Category",
    "Pitch For",
    "Keyword",
]

OUTPUT_COLUMNS = QUERY_INFO_COLUMNS + [
    "business_name",
    "mobile_number",
    "address",
    "website",
    "listing_url",
    "rating",
    "open_close_time",
    "Description",
]

# Added to each input file if missing.
# The API_* / Response_* columns describe the lead_excel_upload call.
TRACKING_COLUMNS = [
    "Start_Date/Time",
    "End_Date/Time",
    "Upload_Status",
    "API_Hit_Time",
    "API_Response_Time(s)",
    "Response_Status",
    "Response_Message",
    "Response_Body",
]
UPLOAD_COLUMNS = TRACKING_COLUMNS[2:]

# Column headers of the CRM lead upload template (in order).
LEAD_TEMPLATE_COLUMNS = [
    "First Name*",
    "Last Name",
    "Email",
    "Mobile Number",
    "Secondary Mobile Number",
    "Fax",
    "Website",
    "Lead Source",
    "Lead Status",
    "Company Name",
    "Industry",
    "No of Employees",
    "Street",
    "City",
    "State",
    "Country",
    "Zip Code",
    "Description",
]

# The CRM accepts only full country names ("Invalid Country: UAE").
COUNTRY_NAMES = {
    "UAE": "United Arab Emirates",
    "KSA": "Saudi Arabia",
    "UK": "United Kingdom",
    "USA": "United States",
    "US": "United States",
}

# Set by either worker to stop both (error, Ctrl+C).
stop_event = threading.Event()
# Set when a scrape completes, so the uploader hits the API right away.
upload_wake = threading.Event()
errors = []
# The scraper and the uploader both read and write the input files.
file_lock = threading.Lock()


class Stopped(Exception):
    pass


def fail(error):
    errors.append(error)
    stop_event.set()
    upload_wake.set()


def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def get_text(card, selector):
    try:
        element = card.locator(selector).first
        if element.count():
            return clean(element.inner_text(timeout=2000))
    except Exception:
        pass
    return ""


def get_attr(card, selector, attribute):
    try:
        element = card.locator(selector).first
        if element.count():
            return clean(element.get_attribute(attribute, timeout=2000))
    except Exception:
        pass
    return ""


def is_hours(text):
    # "Open · Closes 6 pm", "Closed · Opens 8 am Sat", "Open 24 hours".
    return bool(re.search(
        r"\b(?:Opens?|Closed?s?|Open 24 hours)\b", text, re.I
    ))


def get_address(card):
    # In the supplied HTML, address is inside a W4Efsd row.
    # Find the row containing the category/address separator.
    try:
        rows = card.locator(".W4Efsd")
        for i in range(rows.count()):
            text = clean(rows.nth(i).inner_text())
            if "·" in text and re.search(
                r"\d", text
            ) and not is_hours(text):
                # Prefer the segment after the category separator.
                parts = [clean(p) for p in text.split("·")]
                for part in reversed(parts):
                    # Skip phone numbers such as "04 575 1963".
                    if re.fullmatch(r"\+?[\d\s().-]+", part):
                        continue
                    if re.search(r"\d", part) and len(part) > 8:
                        return part
    except Exception:
        pass

    return ""


def get_phone(card):
    phone = get_text(card, PHONE_SELECTOR)
    if phone:
        return phone

    # Fallback: find a phone-like value in card text.
    try:
        text = card.inner_text()
        match = re.search(
            r"\+\d[\d\s().-]{7,}\d",
            text
        )
        if match:
            return clean(match.group())
    except Exception:
        pass

    return ""


def get_rating(card):
    rating = get_text(card, RATING_SELECTOR)
    if rating:
        return rating

    # Fallback: star image label such as "4.5 stars 120 Reviews".
    label = get_attr(card, 'span[role="img"][aria-label]', "aria-label")
    match = re.search(r"(\d+(?:[.,]\d)?)\s*star", label, re.I)
    return match.group(1) if match else ""


def get_hours(card):
    # Cards show a status line such as "Open · Closes 6 pm",
    # "Closed · Opens 9 am Sat" or "Open 24 hours".
    try:
        for line in card.inner_text().splitlines():
            line = clean(line)
            if re.match(
                r"^(Open|Closed|Closes|Opens|Temporarily closed|"
                r"Permanently closed)\b",
                line,
                re.I
            ):
                # Drop a trailing phone number on the same line.
                parts = [
                    clean(p) for p in re.split(r"[·⋅]", line)
                    if clean(p)
                    and not re.fullmatch(r"\+?[\d\s().-]{7,}", clean(p))
                ]
                return " · ".join(parts)
    except Exception:
        pass

    return ""


def extract_card(card, page_url):
    name = get_attr(card, NAME_SELECTOR, "aria-label")
    if not name:
        name = get_text(card, ".qBF1Pd")

    website = get_attr(card, WEBSITE_SELECTOR, "href")
    if website:
        website = urljoin(page_url, website)

    listing_url = get_attr(card, NAME_SELECTOR, "href")
    if listing_url:
        listing_url = urljoin(page_url, listing_url)

    return {
        "business_name": name,
        "mobile_number": get_phone(card),
        "address": get_address(card),
        "website": website,
        "listing_url": listing_url,
        "rating": get_rating(card),
        "open_close_time": get_hours(card),
    }


def scrape_query(page, query):
    records = []
    seen = set()
    idle_scrolls = 0

    page.goto(
        SEARCH_URL.format(quote_plus(query)),
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("If prompted, complete any normal browser steps manually.")
    page.wait_for_timeout(5000)

    cards = page.locator(CARD_SELECTOR)

    # The supplied page uses a scrollable results pane.
    # Inspect the page and update this selector if needed.
    results_panel = page.locator('div[role="feed"]').first

    while len(records) < MAX_LISTINGS:
        if stop_event.is_set():
            raise Stopped()
        card_count = cards.count()
        new_this_round = 0

        for i in range(card_count):
            card = cards.nth(i)
            row = extract_card(card, page.url)

            name = row["business_name"]
            listing_url = row["listing_url"]
            key = listing_url or (
                f'{name}|{row["address"]}'.lower()
            )

            if not name or key in seen:
                continue

            seen.add(key)
            records.append(row)
            new_this_round += 1

            print(
                f'{len(records)}. {name} | '
                f'{row["mobile_number"]} | '
                f'{row["rating"]} | '
                f'{row["open_close_time"]}'
            )

            if len(records) >= MAX_LISTINGS:
                break

        print(
            f"Listings collected: {len(records)}"
        )

        if len(records) >= MAX_LISTINGS:
            break

        if new_this_round == 0:
            idle_scrolls += 1
        else:
            idle_scrolls = 0

        if idle_scrolls >= MAX_IDLE_SCROLLS:
            print("No new listings found after scrolling.")
            break

        # Scroll the results panel to load additional cards.
        try:
            if results_panel.count():
                results_panel.evaluate(
                    "(el) => el.scrollBy(0, Math.max("
                    "500, el.clientHeight * 0.8))"
                )
            else:
                page.mouse.wheel(0, 700)
        except Exception:
            page.mouse.wheel(0, 700)

        page.wait_for_timeout(
            SCROLL_PAUSE * 1000
        )

    return records


def output_path(query):
    # Named after the search query; numbered if the file already exists.
    name = re.sub(r"[^A-Za-z0-9]+", "_", query).strip("_")
    path = os.path.join(OUTPUT_DIR, f"{name}.xlsx")
    counter = 2
    while os.path.exists(path):
        path = os.path.join(OUTPUT_DIR, f"{name}_{counter}.xlsx")
        counter += 1
    return path


def input_files():
    # Re-listed on every check so files added to the folder are picked up.
    # "~$name.xlsx" is Excel's lock file, not a query file.
    return sorted(
        os.path.join(INPUT_DIR, name) for name in os.listdir(INPUT_DIR)
        if name.lower().endswith(".xlsx") and not name.startswith("~$")
    )


def is_open(input_file):
    # Excel keeps a "~$name.xlsx" lock file next to an open workbook.
    folder, name = os.path.split(os.path.abspath(input_file))
    return os.path.exists(os.path.join(folder, "~$" + name))


def read_queries():
    # Rows of every input file, file by file in order. Re-read on every
    # check so rows added to a file are picked up. A file open in Excel is
    # skipped until it is closed.
    queries = []
    for input_file in input_files():
        with file_lock:
            if is_open(input_file):
                print(f"Skipping {input_file}: open in Excel.")
                continue
            try:
                workbook = load_workbook(input_file, read_only=True)
            except PermissionError:
                print(f"Skipping {input_file}: locked by another program.")
                continue
            rows = list(workbook.active.iter_rows(values_only=True))
            workbook.close()

        headers = [text(header) for header in rows[0]] if rows else []
        if QUERY_COLUMN not in headers:
            print(f'Skipping {input_file}: no "{QUERY_COLUMN}" column.')
            continue
        queries += [
            {"_file": input_file, "_row": number, **dict(zip(headers, values))}
            for number, values in enumerate(rows[1:], start=2)
        ]
    return queries


def row_key(row):
    # The same query may appear in more than one input file.
    return row["_file"], text(row.get(QUERY_COLUMN))


def update_row(row, values):
    # Write only this row's cells, saved after every status change so
    # progress survives a crash and other rows are left as they are.
    # Saving over a workbook open in Excel would lose the user's edits, so
    # the save waits until the file is closed instead of stopping the run.
    input_file = row["_file"]
    waiting = False
    while True:
        with file_lock:
            if not is_open(input_file):
                try:
                    write_row(row, values)
                    return
                except PermissionError:
                    pass  # Locked by another program; wait as for Excel.
        if not waiting:
            print(
                f"{input_file} is open or locked; close it to save "
                f"{text(row[QUERY_COLUMN])}. Checking every "
                f"{FILE_OPEN_RECHECK} s."
            )
            waiting = True
        if stop_event.wait(FILE_OPEN_RECHECK):
            raise Stopped()


def write_row(row, values):
    input_file = row["_file"]
    workbook = load_workbook(input_file)
    sheet = workbook.active
    headers = [text(cell.value) for cell in sheet[1]]
    query_column = headers.index(QUERY_COLUMN) + 1
    query = text(row[QUERY_COLUMN])

    # Rows may have been inserted above since it was read.
    number = row["_row"]
    if text(sheet.cell(number, query_column).value) != query:
        number = next((
            r for r in range(2, sheet.max_row + 1)
            if text(sheet.cell(r, query_column).value) == query
        ), None)
        if number is None:
            print(f"Row no longer in {input_file}: {query}")
            return
        row["_row"] = number

    for column, value in values.items():
        if column not in headers:
            headers.append(column)
            sheet.cell(1, len(headers)).value = column
        cell = sheet.cell(number, headers.index(column) + 1)
        cell.value = value
        if column == "File_name" and value:
            # Clickable; links are relative to the input file.
            cell.hyperlink = os.path.relpath(
                str(value), os.path.dirname(input_file)
            )
            cell.style = "Hyperlink"

    workbook.save(input_file)


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def text(value):
    # Excel cells read back as NaN when empty.
    return "" if pd.isna(value) else clean(str(value))


def find_key(data, key):
    # Search a JSON response at any depth, e.g. {"data": {"token": ...}}.
    if isinstance(data, dict):
        if data.get(key):
            return data[key]
        values = data.values()
    elif isinstance(data, list):
        values = data
    else:
        return None

    for value in values:
        found = find_key(value, key)
        if found:
            return found
    return None


def fill_placeholders(payload, values):
    if isinstance(payload, dict):
        return {k: fill_placeholders(v, values) for k, v in payload.items()}
    if isinstance(payload, str):
        for name, value in values.items():
            payload = payload.replace("{{" + name + "}}", str(value))
    return payload


def is_google_redirect(url):
    return bool(re.match(
        r"https?://(www\.)?google\.[a-z.]+/(aclk|url)\b", url, re.I
    ))


def description_items(row, query_row):
    # (label, value) pairs for the Description; empty values are skipped.
    # State / City / Pitch For / Keyword / search query come from the
    # input file.
    items = [
        ("State", query_row.get("State")),
        ("City", query_row.get("City / Area")),
        ("Rating", row.get("rating")),
        ("Hours", row.get("open_close_time")),
        ("Google Map", row.get("listing_url")),
        ("Pitch For", query_row.get("Pitch For")),
        ("Keyword", query_row.get("Keyword")),
        ("Google Maps Search Query", query_row.get(QUERY_COLUMN)),
    ]
    return [(label, text(value)) for label, value in items if text(value)]


def description_lines(row, query_row, escape):
    # The Google Map line is an HTML link in both the output file and
    # the CRM upload, instead of the raw listing URL.
    lines = []
    for label, value in description_items(row, query_row):
        if label == "Google Map":
            value = (
                f'<a href="{html.escape(value)}" target="_blank">'
                "Open in Google Maps</a>"
            )
        elif escape:
            value = html.escape(value)
        lines.append(f"{label}: {value}")
    return lines


def description(row, query_row):
    # Output file: one item per line.
    return "\n".join(description_lines(row, query_row, escape=False))


def description_html(row, query_row):
    # The CRM renders the Description as HTML: one item per line.
    return "<br>".join(description_lines(row, query_row, escape=True))


def format_descriptions(file_name):
    # Wrap the Description so each item shows on its own line.
    workbook = load_workbook(file_name)
    sheet = workbook.active
    headers = [cell.value for cell in sheet[1]]
    desc_col = headers.index("Description") + 1
    sheet.column_dimensions[
        sheet.cell(row=1, column=desc_col).column_letter
    ].width = 80
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=desc_col).alignment = Alignment(
            wrap_text=True, vertical="top"
        )
    workbook.save(file_name)


def lead_rows(file_name, query_row):
    # The API reads only the CRM lead template columns; anything else
    # is ignored and the file is rejected as "Empty excel".
    scraped = pd.read_excel(file_name, dtype=object).fillna("")
    leads = pd.DataFrame({
        "First Name*": scraped["business_name"],
        "Mobile Number": scraped["mobile_number"],
        # Sponsored listings link through Google ad redirects (500+ chars),
        # which crash the API with "Something went wrong".
        "Website": scraped["website"].map(
            lambda url: "" if is_google_redirect(text(url)) else text(url)
        ),
        "Company Name": scraped["business_name"],
        "Industry": scraped["Category"],
        # Older output files may have the phone number or the opening
        # hours ("Opens 8 am Sat") in "address".
        "Street": scraped["address"].where(
            (scraped["address"] != scraped["mobile_number"])
            & ~scraped["address"].map(lambda value: is_hours(text(value))),
            ""
        ),
        "City": scraped["City / Area"],
        "State": text(query_row.get("State")),
        "Country": scraped["Country"].map(
            lambda country: COUNTRY_NAMES.get(
                text(country).upper(), text(country)
            )
        ),
        "Description": scraped.apply(
            description_html, axis=1, args=(query_row,)
        ),
    }).reindex(columns=LEAD_TEMPLATE_COLUMNS, fill_value="")
    return leads


def lead_template_excel(leads):
    buffer = io.BytesIO()
    leads.to_excel(buffer, sheet_name="Leads", index=False)
    return buffer.getvalue()


def upload_report(session, data):
    # The upload response carries a per-row report ("Upload Status"
    # column) as "base64" and/or a download "url".
    try:
        if data.get("base64"):
            content = base64.b64decode(str(data["base64"]).split(",")[-1])
        elif data.get("url"):
            response = session.get(data["url"], timeout=60)
            response.raise_for_status()
            content = response.content
        else:
            return None
        return pd.read_excel(io.BytesIO(content), dtype=object).fillna("")
    except Exception as error:
        print(f"[API] Could not read upload report: {error}")
        return None


def rejected_places(report, leads):
    # Returns ({row: ["City", "State"]} for rows rejected only for a City /
    # State missing from the CRM's master list, number of other failed rows).
    if report is None or "Upload Status" not in report or (
        len(report) != len(leads)
    ):
        return {}, None
    fixes, others = {}, 0
    for position, status in enumerate(report["Upload Status"]):
        errors = [clean(part) for part in text(status).split(",")]
        errors = [error for error in errors if error]
        places = [
            match.group(1) for match in (
                re.match(r"Invalid (City|State):", error, re.I)
                for error in errors
            ) if match
        ]
        if errors and len(places) == len(errors):
            fixes[leads.index[position]] = [p.title() for p in places]
        elif errors and not re.search(r"success|uploaded|created", status, re.I):
            others += 1
    return fixes, others


def without_places(leads, fixes):
    # Blank the rejected City / State; the Description already has them.
    retry = leads.loc[list(fixes)].copy()
    for index, fields in fixes.items():
        for field in fields:
            retry.at[index, field] = ""
    return retry


class LeadUploader:
    def __init__(self):
        with open(API_CONFIG_FILE, encoding="utf-8") as file:
            self.config = json.load(file)
        # One session so the captcha's cookie is sent with the login.
        self.session = requests.Session()
        self.token = None

    def login(self):
        response = self.session.get(self.config["captcha_url"], timeout=30)
        response.raise_for_status()
        captcha = response.json()
        code = captcha.get("captcha_code")
        if not code:
            raise RuntimeError(f"No captcha_code in response: {captcha}")

        payload = fill_placeholders(self.config["login_payload"], {
            "code": code,
            "sessionId": captcha.get("sessionId", ""),
        })
        response = self.session.post(
            self.config["login_url"], json=payload, timeout=30
        )
        data = response.json() if response.content else {}
        token = find_key(data, "token")
        if not token:
            raise RuntimeError(
                f"Login failed (HTTP {response.status_code}): "
                f"{data.get('message') or response.text[:200]}"
            )
        # Keep only the token itself, without any "Bearer " prefix.
        self.token = re.sub(r"^Bearer\s+", "", str(token))
        print("Logged in to lead API.")

    def upload(self, leads):
        if not self.token:
            self.login()

        # Same format as the web app: a data URL (FileReader.readAsDataURL)
        # in "upload_file".
        encoded = base64.b64encode(lead_template_excel(leads)).decode()
        payload = {"upload_file": f"data:{XLSX_MIME};base64,{encoded}"}

        response = self.post_upload(payload)
        if response.status_code == 401:
            # Token expired: log in again once and retry.
            self.login()
            response = self.post_upload(payload)
        return response

    def post_upload(self, payload):
        return self.session.post(
            self.config["upload_url"],
            json=payload,
            headers={"Authorization": f"Bearer {self.token}"},
            timeout=120,
        )


def send_leads(uploader, leads):
    response = uploader.upload(leads)
    try:
        data = response.json()
    except ValueError:
        data = {}
    success = response.ok and not (
        isinstance(data, dict) and data.get("success") is False
    )
    return response, data, success


def upload_row(uploader, row):
    # Returns the API_* / Response_* cells to write for this row.
    file_name = text(row.get("File_name"))
    result = {"API_Hit_Time": now()}
    started = time.perf_counter()

    try:
        if not os.path.exists(file_name):
            raise FileNotFoundError(f"Output file not found: {file_name}")

        leads = lead_rows(file_name, row)
        response, data, success = send_leads(uploader, leads)
        note = ""
        if not success and isinstance(data, dict):
            fixes, others = rejected_places(
                upload_report(uploader.session, data), leads
            )
            if fixes:
                print(
                    f"[API] {len(fixes)} rows rejected for City/State; "
                    "uploading them again without those fields."
                )
                response, data, success = send_leads(
                    uploader, without_places(leads, fixes)
                )
                note = f"Re-sent {len(fixes)} rows without invalid City/State"
                # Rows rejected for other reasons still need a fix.
                if others:
                    success = False
                    note += f"; {others} rows failed for other reasons"
                note += ": "

        message = (
            data.get("message") if isinstance(data, dict) else None
        ) or response.reason or response.text[:200]
        result["Response_Status"] = response.status_code
        result["Response_Message"] = note + clean(str(message))
        # Full response, e.g. the field-level errors behind
        # "Validation error". Excel cells hold at most 32,767 chars.
        # The "base64" field is the whole report file; keep its "url".
        body = json.dumps(
            {k: v for k, v in data.items() if k != "base64"}
        ) if isinstance(data, dict) and data else response.text
        result["Response_Body"] = body[:32000]
        print(f"[API] Upload response body: {body[:2000]}")
    except Exception as error:
        success = False
        result["Response_Status"] = "Error"
        result["Response_Message"] = clean(
            str(error).splitlines()[0] if str(error)
            else type(error).__name__
        )

    result["API_Response_Time(s)"] = round(time.perf_counter() - started, 2)
    result["Upload_Status"] = "Uploaded" if success else "Failed"
    print(
        f'[API] Upload {result["Upload_Status"]}: '
        f'{text(row.get(QUERY_COLUMN))} -> '
        f'{result["Response_Status"]} {result["Response_Message"]}'
    )
    return result


def in_upload_window():
    hour = datetime.now(IST).hour
    if UPLOAD_START_HOUR <= UPLOAD_END_HOUR:
        return UPLOAD_START_HOUR <= hour < UPLOAD_END_HOUR
    # The window runs past midnight.
    return hour >= UPLOAD_START_HOUR or hour < UPLOAD_END_HOUR


def seconds_until_upload_window():
    current = datetime.now(IST)
    start = current.replace(
        hour=UPLOAD_START_HOUR, minute=0, second=0, microsecond=0
    )
    if start <= current:
        start += timedelta(days=1)
    return (start - current).total_seconds()


def upload_worker():
    # Uploads every scraped row not yet uploaded, only inside the upload
    # window; failed uploads are hit again every UPLOAD_RETRY_DELAY
    # seconds until they succeed.
    try:
        uploader = LeadUploader()
        last_try = {}
        while not stop_event.is_set():
            if not in_upload_window():
                wait = seconds_until_upload_window()
                print(
                    f"[API] Outside upload window "
                    f"({UPLOAD_START_HOUR:02d}:00-{UPLOAD_END_HOUR:02d}:00 "
                    f"IST); uploads resume in {wait / 3600:.1f} h."
                )
                # Only a stop ends this wait early, not a finished scrape.
                stop_event.wait(wait)
                continue
            for row in read_queries():
                if stop_event.is_set() or not in_upload_window():
                    break
                key = row_key(row)
                if (
                    not key[1]
                    or text(row.get("Scraped?")) != "Completed"
                    or text(row.get("Upload_Status")) in (
                        "Uploaded", "No records"
                    )
                    or time.time() - last_try.get(key, 0)
                    < UPLOAD_RETRY_DELAY
                ):
                    continue
                last_try[key] = time.time()
                update_row(row, upload_row(uploader, row))
            # A finished scrape sets upload_wake, so new files are
            # uploaded right away instead of after the full interval.
            upload_wake.wait(POLL_INTERVAL)
            upload_wake.clear()
    except Stopped:
        pass
    except Exception as error:
        fail(error)


def scrape_row(page, row):
    query = text(row[QUERY_COLUMN])
    print(f"\n===== {query} =====")
    update_row(row, {
        "Scraped?": "Started",
        "Start_Date/Time": now(),
        "End_Date/Time": None,
        **{column: None for column in UPLOAD_COLUMNS},
    })

    try:
        records = scrape_query(page, query)
    except Stopped:
        # Left as "Started"; scraped again on the next run.
        raise
    except Exception as error:
        print(f"Failed: {query} -> {error}")
        update_row(row, {"Scraped?": "Failed", "End_Date/Time": now()})
        return False

    df = pd.DataFrame(records)
    if not df.empty:
        df = df.drop_duplicates(subset=["listing_url"], keep="first")
    for column in QUERY_INFO_COLUMNS:
        df[column] = row.get(column)
    if not df.empty:
        df["Description"] = df.apply(description, axis=1, args=(row,))
    df = df.reindex(columns=OUTPUT_COLUMNS)

    file_name = output_path(query)
    df.to_excel(file_name, index=False)
    format_descriptions(file_name)

    # The upload worker picks up "Completed" rows on its own.
    values = {
        "Scraped?": "Completed",
        "Records": len(df),
        "File_name": file_name,
        "End_Date/Time": now(),
    }
    if df.empty:
        values["Upload_Status"] = "No records"
    update_row(row, values)
    upload_wake.set()
    print(f"Export complete: {len(df)} records saved to {file_name}")
    return True


def scrape_worker(playwright):
    # Scrapes every row not yet "Completed", including rows and files added
    # while running. Rows come file by file, so the first file's pending
    # rows are all done before the next file is started. Any status other
    # than "Completed" (empty, "Started" from an interrupted run, "Failed")
    # is scraped; a query that fails is retried after SCRAPE_RETRY_DELAY.
    # The browser is opened only when a row needs scraping and closed
    # again once none are left.
    failed = {}  # row key -> time it last failed
    current_file = None
    browser = page = None
    try:
        while not stop_event.is_set():
            row = next((
                row for row in read_queries()
                if text(row.get(QUERY_COLUMN))
                and text(row.get("Scraped?")) != "Completed"
                and time.time() - failed.get(row_key(row), 0)
                >= SCRAPE_RETRY_DELAY
            ), None)
            if row is None:
                if browser:
                    print("Nothing left to scrape; closing the browser.")
                    browser.close()
                    browser = page = None
                print(
                    f"No queries to scrape; checking again in "
                    f"{POLL_INTERVAL // 60} min."
                )
                stop_event.wait(POLL_INTERVAL)
                continue
            if browser is None:
                browser = playwright.chromium.launch(headless=True)
                page = browser.new_page(
                    viewport={"width": 1440, "height": 1000}
                )
            if row["_file"] != current_file:
                current_file = row["_file"]
                print(f"\n##### Input file: {current_file} #####")
            if not scrape_row(page, row):
                failed[row_key(row)] = time.time()
    finally:
        if browser:
            browser.close()


def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Scraping (Playwright, this thread) and uploading run side by side.
    uploader_thread = threading.Thread(target=upload_worker, daemon=True)
    uploader_thread.start()
    print(
        f"Running. Watching {INPUT_DIR}/ every {POLL_INTERVAL // 60} min; "
        f"uploading {UPLOAD_START_HOUR:02d}:00-{UPLOAD_END_HOUR:02d}:00 IST; "
        "press Ctrl+C to stop."
    )

    try:
        with sync_playwright() as p:
            scrape_worker(p)
    except Stopped:
        pass
    except KeyboardInterrupt:
        print("\nStopping...")
    except Exception as error:
        fail(error)
    finally:
        stop_event.set()
        upload_wake.set()
        uploader_thread.join()

    if errors:
        print(f"\nERROR - stopped: {errors[0]}")
        sys.exit(1)


if __name__ == "__main__":
    main()
