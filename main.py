import os, re, json, time, hashlib, logging, csv
from datetime import datetime, timezone, date
from urllib.parse import urljoin, urlparse, parse_qs
from zoneinfo import ZoneInfo

from playwright.sync_api import sync_playwright
import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------------
BASE_URL = 'https://ssl.doas.state.ga.us/gpr/'
SEARCH_URL = BASE_URL + 'index?persisted=true'
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '1T_cg_DBR5WG7vrQZuNqV4hP6BC_BY5sMpwx1v7j5IoU')
TRACKER_GID = int(os.getenv('TRACKER_GID', '2125924989'))
REVIEW_GID = int(os.getenv('REVIEW_GID', '1667806659'))
KEYWORD_LOG_TITLE = os.getenv('KEYWORD_LOG_TITLE', 'Keyword Run Log')
MAX_DETAILS = int(os.getenv('MAX_DETAILS_PER_RUN', '200'))
MAX_PAGES_PER_KEYWORD = int(os.getenv('MAX_PAGES_PER_KEYWORD', '10'))
DELAY = float(os.getenv('REQUEST_DELAY_SECONDS', '0.5'))
RETRIES = int(os.getenv('PAGE_RETRIES', '3'))
TZ = ZoneInfo('America/New_York')          # Georgia time, used for "days until deadline"
DEBUG_DIR = 'data/debug'

KEYWORDS = [
    'Electronic Health Record', 'Electronic Medical Record', 'Health Information Exchange', 'Health Information Technology',
    'Healthcare Information System', 'Clinical Information System', 'Hospital Information System', 'Electronic Visit Verification',
    'Medicaid Management Information System', 'Public Health Information System', 'Healthcare Analytics', 'Healthcare Interoperability',
    'Application Development', 'Application Modernization', 'Data Integration', 'API Integration', 'Cloud Migration',
    'Managed Security Services', 'Healthcare Cybersecurity', 'IT Staff Augmentation', 'Artificial Intelligence', 'Machine Learning',
    'Integrated Eligibility'
]

# New columns (appended automatically to the sheets if they do not exist yet).
EXTRA_COLUMNS = ['Failure Reason', 'Missing Fields', 'Last Changed Date', 'Change Summary']

TRACKER_HEADERS = [
    'Opportunity ID', 'Date Found', 'State', 'Agency', 'Opportunity Title', 'Solicitation / Reference No.', 'Opportunity Type', 'Service Category',
    'Official Source URL', 'Agency Site URL', 'Event ID', 'Event Type', 'Purchase Type', 'Category Type', 'Government Type', 'Fiscal Year',
    'Published Date', 'Deadline Date', 'Deadline Time / Time Zone', 'Days Until Deadline', 'Estimated Value (if published)', 'Scope Summary',
    'Eligibility / Mandatory Requirements', 'Required Certifications / Experience', 'Contract Duration / Options', 'Incumbent / Prior Award (verified only)',
    'Healthcare IT Fit (0-25)', 'Eligibility & Compliance (0-20)', 'Customer / Contract Fit (0-15)', 'Delivery Capability (0-15)',
    'Deadline Feasibility (0-10)', 'Commercial Attractiveness (0-10)', 'Information Completeness (0-5)', 'Total Score (0-100)',
    'Mandatory Gate (Pass / Fail / Unknown)', 'Recommendation (Go / Investigate / No-Go)', 'Fit Rationale', 'Risks / Gaps', 'Next Action',
    'Last Verified Date', 'Owner', 'Review Status', 'Reviewer Notes', 'Search Keywords Matched', 'First Seen Date', 'Last Seen Date',
    'Event Status', 'Event Start Date/Time', 'Buyer Contact Name', 'Buyer Contact Phone', 'Buyer Contact Email', 'Buyer Contact Address',
    'Contact Extraction Status', 'Last Updated At', 'Record Hash'
] + EXTRA_COLUMNS

REVIEW_HEADERS = [
    'Opportunity ID', 'Date Added to Review', 'State', 'Agency', 'Opportunity Title', 'Solicitation / Reference No.',
    'Opportunity Type', 'Service Category', 'Official Source URL', 'Agency Site URL', 'Event ID', 'Event Type', 'Event Status',
    'Purchase Type', 'Category Type', 'Government Type', 'Fiscal Year', 'Published Date', 'Deadline Date', 'Deadline Time / Time Zone',
    'Days Until Deadline', 'Estimated Value (if published)', 'Scope Summary', 'Eligibility / Mandatory Requirements',
    'Required Certifications / Experience', 'Contract Duration / Options', 'Incumbent / Prior Award (verified only)',
    'Healthcare IT Fit (0-25)', 'Eligibility & Compliance (0-20)', 'Customer / Contract Fit (0-15)', 'Delivery Capability (0-15)',
    'Deadline Feasibility (0-10)', 'Commercial Attractiveness (0-10)', 'Information Completeness (0-5)', 'Total Score (0-100)',
    'Mandatory Gate (Pass / Fail / Unknown)', 'Recommendation (Go / Investigate / No-Go)', 'Fit Rationale', 'Risks / Gaps',
    'Next Action', 'Last Verified Date', 'Owner', 'Review Status', 'Reviewer Notes', 'Search Keywords Matched', 'First Seen Date',
    'Last Seen Date', 'Event Start Date/Time', 'Buyer Contact Name', 'Buyer Contact Phone', 'Buyer Contact Email',
    'Buyer Contact Address', 'Contact Extraction Status', 'Last Updated At', 'Record Hash'
] + EXTRA_COLUMNS

KEYWORD_LOG_HEADERS = ['Run Date', 'Run Time (UTC)', 'Keyword', 'Status Filter', 'Total Reported By Site',
                       'Result Links Found', 'New Events (not stored before)', 'Pages Scraped', 'Error']

# Fields filled by humans in the sheet: a non-empty sheet value always wins over the scraper.
HUMAN_FIELDS = [
    'Healthcare IT Fit (0-25)', 'Eligibility & Compliance (0-20)', 'Customer / Contract Fit (0-15)', 'Delivery Capability (0-15)',
    'Deadline Feasibility (0-10)', 'Commercial Attractiveness (0-10)', 'Information Completeness (0-5)', 'Total Score (0-100)',
    'Mandatory Gate (Pass / Fail / Unknown)', 'Recommendation (Go / Investigate / No-Go)', 'Fit Rationale', 'Risks / Gaps',
    'Next Action', 'Owner', 'Review Status', 'Reviewer Notes', 'Eligibility / Mandatory Requirements',
    'Required Certifications / Experience', 'Contract Duration / Options', 'Incumbent / Prior Award (verified only)'
]
# Fields compared run-to-run so a buyer/owner edit on the website is detected and written to the sheet.
CHANGE_FIELDS = [
    'Opportunity Title', 'Agency', 'Event Status', 'Published Date', 'Deadline Date', 'Deadline Time / Time Zone',
    'Event Start Date/Time', 'Estimated Value (if published)', 'Scope Summary', 'Agency Site URL',
    'Buyer Contact Name', 'Buyer Contact Phone', 'Buyer Contact Email', 'Buyer Contact Address'
]
CLOSED_STATUSES = {'closed', 'cancelled', 'canceled', 'awarded'}
VALID_STATUSES = {'open', 'closed', 'cancelled', 'canceled', 'awarded', 'evaluated', 'pending', 'draft'}
LABEL_STOPS = ['Event Type', 'Event Status', 'Purchase Type', 'Category Type', 'Government Type', 'Fiscal Year',
               'Start Date', 'End Date', 'Ends In', 'Buyer Contact', 'Description', 'Event Details', 'Documents',
               'NIGP Codes', 'Procurement Help Desk', 'Link']

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')


# ----------------------------------------------------------------------------
# SMALL HELPERS
# ----------------------------------------------------------------------------
def clean(s):
    return re.sub(r'\s+', ' ', str(s if s is not None else '')).strip()


def norm(s):
    return re.sub(r'[^a-z0-9]+', '', (s or '').lower())


def today_str():
    return datetime.now(TZ).strftime('%Y-%m-%d')


def now_iso():
    return datetime.now(TZ).isoformat(timespec='seconds')


def url_key(u):
    return clean(u).split('#')[0]


def hash_record(row):
    return hashlib.sha256('|'.join(clean(row.get(k, '')) for k in CHANGE_FIELDS).encode()).hexdigest()[:20]


MONTHS = {m: i for i, m in enumerate(['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'], 1)}


def parse_date_time(s):
    """Return (date|None, 'h:mm AM' string, tz string) from text like 'Nov 5, 2026 2:00 PM EST' or '11/05/2026 14:00'."""
    s = clean(s)
    d = None
    m = re.search(r'\b(\d{1,2})/(\d{1,2})/(\d{4})\b', s)
    if m:
        try: d = date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
        except ValueError: pass
    if not d:
        m = re.search(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b', s, re.I)
        if m:
            try: d = date(int(m.group(3)), MONTHS[m.group(1)[:3].lower()], int(m.group(2)))
            except ValueError: pass
    if not d:
        m = re.search(r'\b(\d{4})-(\d{2})-(\d{2})\b', s)
        if m:
            try: d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError: pass
    t = re.search(r'\b(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?\b', s, re.I)
    time_s = clean(f"{t.group(1)}:{t.group(2)} {t.group(3) or ''}").upper() if t else ''
    tz = re.search(r'\b(ET|EST|EDT|CT|CST|CDT|MT|MST|MDT|PT|PST|PDT|UTC)\b', s)
    return d, time_s, (tz.group(1) if tz else '')


def with_retries(fn, what, attempts=RETRIES):
    last = None
    for n in range(1, attempts + 1):
        try:
            return fn()
        except Exception as e:  # noqa
            last = e
            logging.warning('%s failed (attempt %s/%s): %s', what, n, attempts, str(e)[:200])
            time.sleep(1.5 * n)
    raise last


# ----------------------------------------------------------------------------
# GOOGLE SHEETS
# ----------------------------------------------------------------------------
def get_google_sheets():
    raw = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not raw:
        raise RuntimeError('Missing GOOGLE_SERVICE_ACCOUNT_JSON GitHub secret')
    creds = Credentials.from_service_account_info(
        json.loads(raw), scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
    book = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
    tracker = next((w for w in book.worksheets() if w.id == TRACKER_GID), None)
    review = next((w for w in book.worksheets() if w.id == REVIEW_GID), None)
    if tracker is None: raise RuntimeError(f'Tracker worksheet gid {TRACKER_GID} not found')
    if review is None: raise RuntimeError(f'Manual Review worksheet gid {REVIEW_GID} not found')
    try:
        log_ws = book.worksheet(KEYWORD_LOG_TITLE)
    except gspread.WorksheetNotFound:
        log_ws = book.add_worksheet(title=KEYWORD_LOG_TITLE, rows=1000, cols=len(KEYWORD_LOG_HEADERS))
    return tracker, review, log_ws


def ensure_headers(ws, headers):
    """Append missing headers to the right; never reorder/delete existing columns."""
    existing = ws.row_values(1)
    while existing and not clean(existing[-1]):
        existing.pop()
    if not existing:
        if ws.col_count < len(headers): ws.add_cols(len(headers) - ws.col_count)
        ws.update(range_name=f'A1:{rowcol_to_a1(1, len(headers))}', values=[headers], value_input_option='RAW')
        return list(headers)
    have = {clean(h).casefold() for h in existing if clean(h)}
    missing = []
    for h in headers:
        if clean(h).casefold() not in have:
            missing.append(h); have.add(clean(h).casefold())
    if not missing:
        return existing
    start, end = len(existing) + 1, len(existing) + len(missing)
    if ws.col_count < end: ws.add_cols(end - ws.col_count)
    ws.update(range_name=f'{rowcol_to_a1(1, start)}:{rowcol_to_a1(1, end)}', values=[missing], value_input_option='RAW')
    logging.info("Worksheet '%s': added %d new header(s)", ws.title, len(missing))
    return existing + missing


def rows_as_dicts(ws):
    vals = ws.get_all_values()
    if not vals: return [], []
    headers = vals[0]
    out = []
    for i, r in enumerate(vals[1:], start=2):
        if not any(clean(x) for x in r): continue
        d = {headers[c]: (r[c] if c < len(r) else '') for c in range(len(headers)) if clean(headers[c])}
        d['_row'] = i
        out.append(d)
    return headers, out


def build_index(existing):
    idx = {}
    for old in existing:
        for k in ('Opportunity ID', 'Event ID', 'Solicitation / Reference No.'):
            v = clean(old.get(k, ''))
            if v and v.lower() not in {'event', 'announcement', 'untitled'}:
                idx.setdefault(v, old['_row'])
        u = url_key(old.get('Official Source URL', ''))
        if u: idx.setdefault('URL:' + u, old['_row'])
    return idx


def upsert_sheet(ws, headers, records):
    """Batched upsert (few API calls instead of one per row)."""
    _, existing = rows_as_dicts(ws)
    idx = build_index(existing)
    updates, appends = [], []
    for rec in records:
        row = (idx.get(clean(rec.get('Opportunity ID', ''))) or idx.get(clean(rec.get('Event ID', '')))
               or idx.get('URL:' + url_key(rec.get('Official Source URL', ''))))
        values = [clean(rec.get(h, '')) if not isinstance(rec.get(h, ''), (int, float)) else rec.get(h) for h in headers]
        if row:
            updates.append({'range': f'A{row}:{rowcol_to_a1(row, len(headers))}', 'values': [values]})
        else:
            appends.append(values)
    for i in range(0, len(updates), 100):
        with_retries(lambda c=updates[i:i + 100]: ws.batch_update(c, value_input_option='RAW'), 'batch_update')
        time.sleep(1)
    for i in range(0, len(appends), 200):
        with_retries(lambda c=appends[i:i + 200]: ws.append_rows(c, value_input_option='RAW'), 'append_rows')
        time.sleep(1)
    return len(appends), len(updates)


def remove_matching_rows(ws, moved_records):
    """Delete rows (one batched request) that now belong to the other worksheet."""
    if not moved_records: return 0
    ids, urls = set(), set()
    for r in moved_records:
        for k in ('Event ID', 'Opportunity ID'):
            if clean(r.get(k, '')): ids.add(clean(r[k]))
        if url_key(r.get('Official Source URL', '')): urls.add(url_key(r['Official Source URL']))
    _, rows = rows_as_dicts(ws)
    targets = [x['_row'] for x in rows
               if clean(x.get('Event ID', '')) in ids or clean(x.get('Opportunity ID', '')) in ids
               or url_key(x.get('Official Source URL', '')) in urls]
    targets = [t for t in set(targets)]
    if targets:
        reqs = [{'deleteDimension': {'range': {'sheetId': ws.id, 'dimension': 'ROWS', 'startIndex': r - 1, 'endIndex': r}}}
                for r in sorted(targets, reverse=True)]
        ws.spreadsheet.batch_update({'requests': reqs})
    return len(targets)


# ----------------------------------------------------------------------------
# SEARCH PAGE
# ----------------------------------------------------------------------------
def parse_total_reported(text):
    for pat in (r'of\s+([\d,]+)\s+(?:entries|results|records|events|opportunities)',
                r'([\d,]+)\s+(?:entries|results|records|events|opportunities)\s+found',
                r'(?:total|results?)\s*(?:records|events|found)?\s*:\s*([\d,]+)'):
        m = re.search(pat, text, re.I)
        if m: return int(m.group(1).replace(',', ''))
    return ''


def collect_links(page, keyword, found):
    hrefs = page.eval_on_selector_all('a[href]', 'els => els.map(e => e.getAttribute("href"))')
    n = 0
    for href in hrefs:
        low = (href or '').lower()
        if 'esourcenumber=' not in low and 'eventdetails' not in low:
            continue
        url = urljoin(page.url, href)
        url = re.sub(r'^https://pur\.doas\.ga\.gov/gpr/', BASE_URL, url)
        key = parse_qs(urlparse(url).query).get('eSourceNumber', [url])[0]
        found[key] = {'url': url, 'keyword': keyword}
        n += 1
    if n == 0:
        # Fallback: results rendered without anchors -> build detail URL from visible event IDs.
        try:
            txt = clean(page.locator('body').inner_text(timeout=5000))
            for eid in set(re.findall(r'\b[A-Z]{2}-[A-Z0-9-]{8,}\b', txt)):
                found.setdefault(eid, {'url': f'{BASE_URL}eventDetails?eSourceNumber={eid}', 'keyword': keyword})
        except Exception:
            pass
    return n


def scrape_search_keyword(page, keyword):
    def load():
        page.goto(SEARCH_URL, wait_until='domcontentloaded', timeout=60000)
        page.wait_for_timeout(1500)
    with_retries(load, f'search page for {keyword}')
    body = ''
    try: body = clean(page.locator('body').inner_text(timeout=5000))[:1200]
    except Exception: pass
    if '/unsupported' in page.url.lower() or 'unsupported browser' in body.lower():
        raise RuntimeError(f'GPR returned unsupported-browser page at {page.url}: {body[:300]}')

    field = page.get_by_placeholder(re.compile('Enter Event ID or Title', re.I)).first
    if field.count() == 0: field = page.locator('input[type="text"]').first
    field.fill(keyword)

    # Status filter = OPEN; also raise page size to the largest option if a page-size dropdown exists.
    for sel in page.locator('select').all():
        try:
            opts = [clean(x) for x in sel.locator('option').all_inner_texts()]
            if any(x.upper() == 'OPEN' for x in opts):
                sel.select_option(label=re.compile('^OPEN$', re.I))
            elif opts and all(re.fullmatch(r'\d+', o) for o in opts):
                sel.select_option(label=max(opts, key=int))
        except Exception:
            pass

    btn = page.get_by_role('button', name=re.compile('^Search$', re.I)).first
    if btn.count() == 0: btn = page.locator('input[type="submit"]').first
    btn.click(timeout=15000)
    try: page.wait_for_load_state('networkidle', timeout=20000)
    except Exception: page.wait_for_timeout(1500)

    found, pages, total = {}, 0, ''
    seen_signatures = set()
    for _ in range(MAX_PAGES_PER_KEYWORD):
        try: total = total or parse_total_reported(clean(page.locator('body').inner_text(timeout=5000)))
        except Exception: pass
        before = set(found)
        page_found = {}
        collect_links(page, keyword, page_found)
        sig = tuple(sorted(page_found))
        if not sig or sig in seen_signatures:     # empty page or pagination stuck on same page
            break
        seen_signatures.add(sig)
        found.update(page_found)
        pages += 1
        nxt = page.get_by_role('link', name=re.compile(r'^(Next|›|»)$', re.I)).last
        if nxt.count() == 0:
            nxt = page.locator('a[aria-label*="Next" i], button[aria-label*="Next" i], li.next a').last
        if nxt.count() == 0 or not nxt.is_visible(): break
        try:
            cls = ((nxt.get_attribute('class') or '') + ' ' + (nxt.locator('xpath=..').get_attribute('class') or '')).lower()
            if 'disabled' in cls or (nxt.get_attribute('aria-disabled') or '').lower() == 'true': break
            nxt.click(timeout=5000)
            page.wait_for_timeout(900)
        except Exception:
            break
        if set(found) == before and pages > 1: break
    return found, {'total_reported': total, 'links_found': len(found), 'pages': pages}


# ----------------------------------------------------------------------------
# DETAIL PAGE
# ----------------------------------------------------------------------------
LABELS = {norm(x) for x in [
    'event id', 'event number', 'esource number', 'event title', 'opportunity title', 'solicitation number', 'reference number',
    'agency', 'agency name', 'entity', 'entity name', 'government entity', 'event type', 'event status', 'purchase type',
    'category type', 'government type', 'fiscal year', 'start date', 'end date', 'closing date', 'deadline date', 'description',
    'event description', 'buyer contact', 'contact name', 'email', 'contact email', 'phone', 'telephone', 'address',
    'estimated value', 'posted date', 'publish date', 'created date', 'date posted', 'ends in', 'link']}


def label_value_map(page):
    pairs = {}

    def add(k, v):
        k, v = clean(k).rstrip(':').strip(), clean(v)
        if not k or not v or norm(k) not in LABELS or norm(v) in LABELS: return
        pairs.setdefault(k, v)

    for tr in page.locator('tr').all():
        try:
            cells = [clean(x).rstrip(':') for x in tr.locator('th,td').all_inner_texts()]
            i = 0
            while i + 1 < len(cells):
                if norm(cells[i]) in LABELS: add(cells[i], cells[i + 1]); i += 2
                else: i += 1
        except Exception: pass
    for dt in page.locator('dt').all():
        try: add(dt.inner_text(), dt.locator('xpath=following-sibling::*[1]').inner_text())
        except Exception: pass
    for lab in page.locator('label').all():
        try:
            target = lab.get_attribute('for')
            add(lab.inner_text(), page.locator('#' + target).input_value() if target else '')
        except Exception: pass
    return pairs


def first_match(mapping, *terms):
    wanted = {norm(t) for t in terms if norm(t)}
    for k, v in mapping.items():
        if norm(k) in wanted and clean(v): return clean(v)
    return ''


def section(text, label, stops=None):
    """Text after `label` up to the next known label."""
    stops = [s for s in (stops or LABEL_STOPS) if s.lower() != label.lower()]
    pat = (rf'(?:^|\s){re.escape(label)}\s*:?\s*(.*?)(?=\s+(?:{"|".join(re.escape(s) for s in sorted(stops, key=len, reverse=True))})\s*:?(?:\s|$)|$)')
    m = re.search(pat, text, re.I)
    return clean(m.group(1)) if m else ''


PHONE_RE = re.compile(r'(?<!\d)(?:\+?1[-. ]?)?\(?\d{3}\)?[-. ]?\d{3}[-. ]\d{4}(?:\s*(?:ext\.?|x|extension)\s*\d{1,6})?(?!\d)', re.I)
EMAIL_RE = re.compile(r'[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}', re.I)


def parse_contact(page, text, mapping):
    """Buyer name / phone / email / address from mailto:/tel: links first, then the 'Buyer Contact' block."""
    block = section(text, 'Buyer Contact') or first_match(mapping, 'buyercontact', 'contactname', 'buyername')
    email = phone = ''
    try:
        for href in page.eval_on_selector_all('a[href^="mailto:" i]', 'els => els.map(e => e.getAttribute("href"))'):
            e = clean(href.split(':', 1)[1].split('?')[0])
            if EMAIL_RE.fullmatch(e): email = e; break
        for href in page.eval_on_selector_all('a[href^="tel:" i]', 'els => els.map(e => e.getAttribute("href"))'):
            p = clean(href.split(':', 1)[1])
            if len(re.sub(r'\D', '', p)) >= 10: phone = p; break
    except Exception:
        pass
    email = email or first_match(mapping, 'email', 'contactemail')
    phone = phone or first_match(mapping, 'phone', 'telephone', 'contactphone')
    if not email:
        m = EMAIL_RE.search(block) or EMAIL_RE.search(text)
        email = m.group(0) if m else ''
    if not phone:
        m = PHONE_RE.search(block) or PHONE_RE.search(text)
        phone = clean(m.group(0)) if m else ''
    if re.search(r'555[-. ]?555[-. ]?\d{4}', phone) or re.sub(r'\D', '', phone) in {'0000000000', '5555555555'}:
        phone = ''
    # name & address from the block with phone/email removed
    rest = block
    for tok in (email, phone):
        if tok: rest = rest.replace(tok, ' ')
    rest = clean(re.sub(r'\b(?:Phone|Tel(?:ephone)?|E-?mail|Fax|Contact|Address)\s*:', ' | ', rest, flags=re.I)).strip(' |')
    name = address = ''
    m = re.match(r"([A-Za-z][A-Za-z.'\-, ]{1,80}?)(?=\s+\d|\s*\||\s*$)", rest)
    if m:
        name = clean(m.group(1)).strip(' ,')
        address = clean(rest[m.end():]).strip(' |,')
    address = first_match(mapping, 'address', 'buyeraddress', 'contactaddress') or address
    if len(address) > 250 or re.search(r'\b(Event Type|Event Status|Purchase Type)\b', address): address = ''
    if name and EMAIL_RE.search(name): name = ''
    return {'name': name, 'phone': phone, 'email': email, 'address': address}


def extract_event_id(url, mapping, text):
    q = parse_qs(urlparse(url).query)
    sid = clean((q.get('eSourceNumber') or q.get('eventId') or [''])[0])
    if sid: return sid
    for k, v in mapping.items():
        if norm(k) in {'eventid', 'eventnumber', 'esourcenumber'}:
            m = re.search(r'\b[A-Z]{2}-[A-Z0-9-]{8,}\b', clean(v), re.I)
            if m: return m.group(0)
    m = re.search(r'\b[A-Z]{2}-[A-Z0-9-]{8,}\b', text, re.I)
    return m.group(0) if m else ''


def extract_title(page, text, mapping):
    for sel in ['h1', 'h2', 'h3', '.page-title', '.event-title']:
        try:
            for c in page.locator(sel).all_inner_texts():
                c = clean(c)
                if re.match(r'^(RFP|RFQ|RFI|ITB|IFB|RFx|RFB|SWC|IRFP)\s*[-#:]?\s*\w+', c, re.I) and len(c) >= 10:
                    return c
        except Exception: pass
    m = re.search(r'((?:RFP|RFQ|RFI|ITB|IFB|RFx|RFB)\s*[-#]?\s*[A-Z0-9][A-Z0-9-]*\s*[-:–—]\s*.{4,220}?)'
                  r'(?=\s+(?:Event Type|Event Status|Purchase Type|Category Type|Government Type|Fiscal Year|Start Date|End Date|Ends In|Event Details|Documents)\b|$)', text, re.I)
    if m: return clean(m.group(1))
    return first_match(mapping, 'eventtitle', 'opportunitytitle', 'solicitationtitle', 'title')


def extract_agency(text, mapping):
    bad = re.compile(r'\b(Event Type|Event Status|Purchase Type|Category Type|Government Type|Fiscal Year|Link)\b', re.I)
    cand = first_match(mapping, 'agency', 'agencyname', 'entity', 'entityname', 'governmententity')
    cand = re.sub(r'^\d{5}\s+', '', cand)
    if cand and len(cand) < 180 and not bad.search(cand) and ' | ' not in cand:
        return cand
    m = re.search(r'\b\d{5}\s+([A-Z][A-Z0-9&.,() /\'-]{2,100}?)(?=\s+(?:Buyer Contact|Description|Event Type|Event Status|Purchase Type|Category Type|Government Type|Fiscal Year|Link|RFP|RFQ|RFI|ITB|IFB|RFx)\b|$)', text)
    if m and not bad.search(m.group(1)): return clean(m.group(1))
    return ''


def extract_agency_site_url(page, title, ref):
    best, best_score = '', 0
    ref_slug = re.sub(r'[^a-z0-9]+', '-', ref.lower()).strip('-')
    words = [w for w in re.findall(r'[a-z0-9]+', title.lower()) if len(w) >= 5]
    try:
        for a in page.locator('a[href]').all():
            href = (a.get_attribute('href') or '').strip()
            if not href or href.startswith(('#', 'javascript:', 'mailto:', 'tel:')) or 'eventdetails' in href.lower(): continue
            absolute = urljoin(page.url, href)
            pu = urlparse(absolute)
            if pu.scheme not in {'http', 'https'}: continue
            if pu.netloc.endswith(('doas.state.ga.us', 'doas.ga.gov')) and 'download' not in absolute.lower(): continue
            if pu.path.rstrip('/').lower() in {'', '/', '/opportunities', '/procurement/opportunities', '/gpr'}: continue
            try: label = clean(a.inner_text(timeout=500))
            except Exception: label = ''
            hay = (label + ' ' + href).lower()
            score = 3 + (10 if ref_slug and ref_slug in re.sub(r'[^a-z0-9]+', '-', hay) else 0) + sum(2 for w in words if w in hay)
            if score > best_score: best, best_score = absolute, score
    except Exception:
        logging.exception('agency-site URL extraction failed')
    return best


def scrape_details(page, url, keywords):
    with_retries(lambda: page.goto(url, wait_until='domcontentloaded', timeout=60000), f'detail {url}')
    try: page.wait_for_load_state('networkidle', timeout=15000)
    except Exception: page.wait_for_timeout(1200)
    text = clean(page.locator('body').inner_text(timeout=15000))
    mapping = label_value_map(page)

    event_id = extract_event_id(url, mapping, text)
    title = extract_title(page, text, mapping)
    agency = extract_agency(text, mapping)

    st = re.search(r'\bEvent Status\s*:?\s*(Open|Closed|Cancelled|Canceled|Awarded|Evaluated|Pending|Draft)\b', text, re.I)
    status = st.group(1).title() if st else first_match(mapping, 'eventstatus').title()
    if status.lower() not in VALID_STATUSES: status = ''

    def field(label, *aliases):
        v = section(text, label)
        if norm(v) in LABELS or len(v) > 150: v = ''
        return v or first_match(mapping, *aliases)

    event_type = field('Event Type', 'eventtype')
    purchase_type = field('Purchase Type', 'purchasetype')
    category_type = field('Category Type', 'categorytype')
    government_type = field('Government Type', 'governmenttype')
    fiscal_year = field('Fiscal Year', 'fiscalyear')

    start_raw = section(text, 'Start Date') or first_match(mapping, 'startdate', 'eventstart')
    end_raw = section(text, 'End Date') or first_match(mapping, 'enddate', 'closingdate', 'duedate', 'deadline', 'deadlinedate')
    pub_raw = first_match(mapping, 'posteddate', 'publishdate', 'createddate', 'dateposted')
    end_d, end_t, end_tz = parse_date_time(end_raw)
    pub_d, _, _ = parse_date_time(pub_raw)
    deadline_date = end_d.isoformat() if end_d else end_raw
    deadline_time = clean(f'{end_t} {end_tz}')
    days_until = ''
    if end_d:
        days_until = str((end_d - datetime.now(TZ).date()).days)
    else:
        dm = re.search(r'Ends In:\s*(-?\d+)\s*days?', text, re.I)
        if dm: days_until = dm.group(1)

    desc = section(text, 'Description') or first_match(mapping, 'description', 'eventdescription', 'scope')
    contact = parse_contact(page, text, mapping)
    value = first_match(mapping, 'estimatedvalue', 'eventvalue', 'totalvalue')

    rm = re.search(r'\b(?:RFP|RFQ|RFI|ITB|IFB|RFx|RFB)\s*[-#]?\s*[A-Z0-9][A-Z0-9-]*', title, re.I)
    solicitation_ref = clean(rm.group(0)) if rm else ''
    agency_site_url = extract_agency_site_url(page, title, solicitation_ref)

    links = f'Event details: {url}#pane-eventDetails\nDocuments: {url}#pane-documents'
    scope = (desc[:3500] + '\n' + links) if desc else links
    found_contact = any(contact.values())

    return {
        'Opportunity ID': event_id or hashlib.sha1(url.encode()).hexdigest()[:12].upper(), 'State': 'Georgia',
        'Agency': agency, 'Opportunity Title': title, 'Solicitation / Reference No.': solicitation_ref,
        'Opportunity Type': event_type or category_type, 'Service Category': category_type,
        'Official Source URL': url, 'Agency Site URL': agency_site_url, 'Event ID': event_id, 'Event Type': event_type,
        'Event Status': status, 'Purchase Type': purchase_type, 'Category Type': category_type, 'Government Type': government_type,
        'Fiscal Year': fiscal_year, 'Published Date': pub_d.isoformat() if pub_d else pub_raw,
        'Deadline Date': deadline_date, 'Deadline Time / Time Zone': deadline_time, 'Days Until Deadline': days_until,
        'Estimated Value (if published)': value, 'Scope Summary': scope,
        'Mandatory Gate (Pass / Fail / Unknown)': 'Unknown', 'Recommendation (Go / Investigate / No-Go)': 'Investigate',
        'Fit Rationale': 'Automated keyword discovery; human review required.',
        'Risks / Gaps': 'Confirm healthcare/IT relevance, requirements, deadline and eligibility in the official event documents.',
        'Next Action': 'Open official event; review documents and complete scoring.',
        'Last Verified Date': today_str(), 'Review Status': 'Needs Review',
        'Search Keywords Matched': '; '.join(sorted({k.strip() for k in keywords if k.strip()})),
        'Event Start Date/Time': start_raw,
        'Buyer Contact Name': contact['name'], 'Buyer Contact Phone': contact['phone'],
        'Buyer Contact Email': contact['email'], 'Buyer Contact Address': contact['address'],
        'Contact Extraction Status': 'Contact found' if found_contact else 'Not found on page',
        'Last Updated At': now_iso(),
        '_page_text': text[:6000],            # only used for debug output, stripped before saving
    }


# ----------------------------------------------------------------------------
# VALIDATION / CHANGE DETECTION
# ----------------------------------------------------------------------------
def validate_record(row):
    """Return (blocking_reasons, optional_missing). Blocking reasons => Manual Review."""
    reasons, optional = [], []
    title, agency = clean(row.get('Opportunity Title')), clean(row.get('Agency'))
    url, eid = clean(row.get('Official Source URL')), clean(row.get('Event ID'))
    if not url or 'eventdetails' not in url.lower() or 'esourcenumber=' not in url.lower():
        reasons.append('Official Source URL missing/invalid')
    if not title or title.lower() in {'event', 'announcement', 'untitled'} or len(title) < 8:
        reasons.append('Opportunity Title missing/invalid')
    if not agency or agency.lower() in {'unknown', 'n/a', 'none', 'agency'}:
        reasons.append('Agency missing')
    if not re.fullmatch(r'[A-Z]{2}-[A-Z0-9-]{8,}', eid, re.I):
        reasons.append('Event ID missing/invalid')
    if not re.search(r'20\d{2}', clean(row.get('Deadline Date'))):
        reasons.append('Deadline Date missing/unparseable')
    if len(re.sub(r'(Event details|Documents):\s*https?://\S+', '', clean(row.get('Scope Summary'))).strip()) < 30:
        reasons.append('Scope/Description missing or too short')
    if clean(row.get('Event Status')).lower() not in VALID_STATUSES:
        reasons.append('Event Status missing/unrecognised')
    for col, label in [('Buyer Contact Name', 'Buyer name'), ('Buyer Contact Phone', 'Buyer phone'),
                       ('Buyer Contact Email', 'Buyer email'), ('Buyer Contact Address', 'Buyer address'),
                       ('Agency Site URL', 'Agency site URL'), ('Published Date', 'Published date'),
                       ('Deadline Time / Time Zone', 'Deadline time')]:
        if not clean(row.get(col)): optional.append(label)
    return reasons, optional


def detect_changes(prior, new):
    changes = []
    for col in CHANGE_FIELDS:
        old, cur = clean(prior.get(col, '')), clean(new.get(col, ''))
        if old != cur:
            changes.append(f"{col}: '{old[:60]}' -> '{cur[:60]}'")
    return changes


def merge_with_prior(row, prior, seen_today):
    """Apply human-entered values, keep good old data if extraction came back blank, and track website changes."""
    # keep previously good scraped values when this run could not extract them (never wipe on a flaky page)
    for col in CHANGE_FIELDS + ['Solicitation / Reference No.', 'Event Type', 'Purchase Type', 'Category Type',
                                 'Government Type', 'Fiscal Year', 'Days Until Deadline']:
        if not clean(row.get(col, '')) and clean(prior.get(col, '')) and col != 'Days Until Deadline':
            row[col] = prior[col]
    for col in HUMAN_FIELDS:
        if clean(prior.get(col, '')): row[col] = prior[col]
    first = clean(prior.get('Date Found')) or clean(prior.get('Date Added to Review')) or today_str()
    row['Date Found'] = first
    row['Date Added to Review'] = clean(prior.get('Date Added to Review')) or clean(prior.get('Date Found')) or today_str()
    row['First Seen Date'] = clean(prior.get('First Seen Date')) or first
    row['Last Seen Date'] = today_str() if seen_today else (clean(prior.get('Last Seen Date')) or today_str())
    # union of keywords
    kws = {k.strip() for k in (clean(prior.get('Search Keywords Matched')) + ';' + clean(row.get('Search Keywords Matched'))).split(';') if k.strip()}
    row['Search Keywords Matched'] = '; '.join(sorted(kws))
    changes = detect_changes(prior, row) if prior else []
    if changes:
        row['Last Changed Date'] = today_str()
        row['Change Summary'] = (today_str() + ': ' + ' | '.join(changes))[:1500]
        logging.info('CHANGED %s: %s', row.get('Event ID'), '; '.join(changes)[:300])
    else:
        row['Last Changed Date'] = clean(prior.get('Last Changed Date'))
        row['Change Summary'] = clean(prior.get('Change Summary'))
    row['Record Hash'] = hash_record(row)
    return row


# ----------------------------------------------------------------------------
# EXPORT
# ----------------------------------------------------------------------------
def export_data(tracker_rows, review_rows, diagnostics):
    os.makedirs('data', exist_ok=True)
    for basename, rows, base_cols in [('georgia_leads', tracker_rows, TRACKER_HEADERS),
                                      ('georgia_manual_review', review_rows, REVIEW_HEADERS)]:
        rows = [{k: v for k, v in r.items() if not k.startswith('_')} for r in rows]
        cols = list(base_cols)                     # copy: never mutate the global header lists
        for r in rows:
            for k in r:
                if k not in cols: cols.append(k)
        with open(f'data/{basename}.json', 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        with open(f'data/{basename}.csv', 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.DictWriter(f, fieldnames=cols, extrasaction='ignore')
            w.writeheader(); w.writerows(rows)
    with open('data/georgia_diagnostics.json', 'w', encoding='utf-8') as f:
        json.dump(diagnostics, f, ensure_ascii=False, indent=2)


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main():
    run_started = datetime.now(timezone.utc)
    tracker, review, log_ws = get_google_sheets()
    tracker_headers = ensure_headers(tracker, TRACKER_HEADERS)
    review_headers = ensure_headers(review, REVIEW_HEADERS)
    ensure_headers(log_ws, KEYWORD_LOG_HEADERS)

    _, old_tracker = rows_as_dicts(tracker)
    _, old_review = rows_as_dicts(review)
    old_all = old_tracker + old_review
    prior_by_id, prior_by_url = {}, {}
    for r in old_all:
        for k in ('Event ID', 'Opportunity ID'):
            if clean(r.get(k)): prior_by_id.setdefault(clean(r[k]), r)
        if url_key(r.get('Official Source URL', '')): prior_by_url.setdefault(url_key(r['Official Source URL']), r)

    def find_prior(event_id, url):
        return prior_by_id.get(clean(event_id)) or prior_by_url.get(url_key(url)) or {}

    found, failures, keyword_stats, log_rows = {}, [], [], []
    detail_failures, skipped_not_open = [], 0
    records = []
    os.makedirs(DEBUG_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={'width': 1440, 'height': 1000}, locale='en-US', timezone_id='America/New_York',
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
            extra_http_headers={'Accept-Language': 'en-US,en;q=0.9'})
        page = context.new_page()

        # ---- 1. search every keyword (Event Status = OPEN) and record how many opportunities each returned
        for i, keyword in enumerate(KEYWORDS, 1):
            err, stats, results = '', {'total_reported': '', 'links_found': 0, 'pages': 0}, {}
            try:
                results, stats = scrape_search_keyword(page, keyword)
                new_cnt = sum(1 for k, v in results.items() if not find_prior(k, v['url']))
                for k, v in results.items():
                    if k in found: found[k]['keywords'].add(keyword)
                    else: found[k] = {'url': v['url'], 'keywords': {keyword}}
            except Exception as e:
                logging.exception('Search failed for keyword: %s', keyword)
                err = str(e)[:300]; new_cnt = 0
                failures.append({'keyword': keyword, 'error': err})
            logging.info('[%s/%s] %s => reported=%s links=%s pages=%s', i, len(KEYWORDS), keyword,
                         stats['total_reported'], stats['links_found'], stats['pages'])
            keyword_stats.append({'keyword': keyword, **stats, 'new_events': new_cnt, 'error': err})
            log_rows.append([today_str(), run_started.strftime('%H:%M:%S'), keyword, 'OPEN', stats['total_reported'],
                             stats['links_found'], new_cnt, stats['pages'], err])
            time.sleep(DELAY)

        # ---- 2. build the detail queue: newly found events + stored events that may still change
        queue = {}
        for key, item in found.items():
            queue[url_key(item['url'])] = (key, item['url'], set(item['keywords']), True)
        for old in old_all:
            u = clean(old.get('Official Source URL', ''))
            if not u or url_key(u) in queue: continue
            if clean(old.get('Event Status', '')).lower() in CLOSED_STATUSES: continue   # final states: stop re-checking
            kws = {x.strip() for x in clean(old.get('Search Keywords Matched', '')).split(';') if x.strip()}
            queue[url_key(u)] = (clean(old.get('Event ID')) or clean(old.get('Opportunity ID')), u, kws, False)
        items = list(queue.values())
        if len(items) > MAX_DETAILS:
            items.sort(key=lambda it: (0 if not find_prior(it[0], it[1]) else 1,
                                       clean(find_prior(it[0], it[1]).get('Last Verified Date', ''))))
            items = items[:MAX_DETAILS]

        # ---- 3. scrape each detail page
        for n, (key, url, kws, seen_today) in enumerate(items, 1):
            try:
                row = scrape_details(page, url, kws)
                prior = find_prior(row.get('Event ID'), url)
                status = clean(row.get('Event Status')).lower()
                if not prior and status in CLOSED_STATUSES:
                    skipped_not_open += 1
                    logging.info('Skipping new non-open event %s (%s)', row.get('Event ID'), status)
                    continue
                row = merge_with_prior(row, prior, seen_today)
                reasons, optional = validate_record(row)
                row['Failure Reason'] = '; '.join(reasons)
                row['Missing Fields'] = ', '.join(optional)
                if reasons:    # save the page text so selectors can be fixed from the Actions artifact
                    safe = re.sub(r'[^A-Za-z0-9_-]+', '_', row.get('Event ID') or row['Opportunity ID'])
                    with open(f'{DEBUG_DIR}/{safe}.txt', 'w', encoding='utf-8') as f:
                        f.write(url + '\n\n' + row.get('_page_text', ''))
                row.pop('_page_text', None)
                row['_reasons'] = reasons
                records.append(row)
                logging.info('Details %s/%s: %s %s', n, len(items), row['Opportunity ID'], 'OK' if not reasons else 'REVIEW: ' + row['Failure Reason'])
            except Exception as e:
                logging.exception('Detail failed for %s', url)
                detail_failures.append({'url': url, 'error': str(e)[:300]})
            time.sleep(DELAY)
        browser.close()

    # ---- 4. split: complete -> Tracker, incomplete -> Manual Review (with the reason)
    lead_records = [r for r in records if not r['_reasons']]
    review_records = [r for r in records if r['_reasons']]
    for r in lead_records: r['Failure Reason'] = ''
    for r in review_records:
        r['Date Added to Review'] = clean(r.get('Date Added to Review')) or today_str()
        if not clean(r.get('Official Source URL')): r['Official Source URL'] = ''
        r['Review Status'] = clean(r.get('Review Status')) or 'Needs Review'
    for r in records: r.pop('_reasons', None)

    if lead_records:
        a, u = upsert_sheet(tracker, tracker_headers, lead_records)
        logging.info('Tracker sync: %s added, %s updated', a, u)
    if review_records:
        a, u = upsert_sheet(review, review_headers, review_records)
        logging.info('Manual review sync: %s added, %s updated', a, u)

    # a record lives in exactly one sheet: remove it from the other one after the write succeeded
    rm_review = remove_matching_rows(review, lead_records)
    rm_tracker = remove_matching_rows(tracker, review_records)
    logging.info('Moved: %s review->tracker, %s tracker->review', rm_review, rm_tracker)

    if log_rows:
        with_retries(lambda: log_ws.append_rows(log_rows, value_input_option='RAW'), 'keyword log')

    total_reported = sum(k['total_reported'] for k in keyword_stats if isinstance(k['total_reported'], int))
    diagnostics = {
        'run_at': run_started.isoformat(), 'status_filter': 'OPEN', 'keywords_total': len(KEYWORDS),
        'keyword_stats': keyword_stats, 'sum_of_reported_results_all_keywords': total_reported,
        'unique_events_discovered': len(found), 'details_attempted': len(items),
        'records_scraped': len(records), 'tracker_records': len(lead_records), 'manual_review_records': len(review_records),
        'changed_records': sum(1 for r in records if clean(r.get('Last Changed Date')) == today_str()),
        'skipped_new_non_open': skipped_not_open, 'moved_review_to_tracker': rm_review, 'moved_tracker_to_review': rm_tracker,
        'search_failures': failures, 'detail_failures': detail_failures,
        'tracker_gid': TRACKER_GID, 'review_gid': REVIEW_GID, 'max_details_per_run': MAX_DETAILS}

    _, tracker_snapshot = rows_as_dicts(tracker)
    _, review_snapshot = rows_as_dicts(review)
    export_data(tracker_snapshot, review_snapshot, diagnostics)
    logging.info('Diagnostics: %s', json.dumps({k: v for k, v in diagnostics.items() if k != 'keyword_stats'}))

    if len(failures) == len(KEYWORDS):
        raise RuntimeError('Every keyword search failed - check selectors / site availability in the Actions log.')
    if not found and not records:
        raise RuntimeError('No event-detail links were found and no stored events could be refreshed.')


if __name__ == '__main__':
    main()
