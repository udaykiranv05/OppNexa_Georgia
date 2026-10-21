import os, re, json, time, hashlib, logging, csv
from datetime import datetime, timezone
from urllib.parse import urljoin, urlparse, parse_qs
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import gspread
from google.oauth2.service_account import Credentials

# Current public GPR pages and event details are served on pur.doas.ga.gov.
BASE_URL = 'https://ssl.doas.state.ga.us/gpr/'
SEARCH_URL = BASE_URL + 'index?persisted=true'
CANONICAL_BASE_URL = 'https://pur.doas.ga.gov/gpr/'
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID', '1T_cg_DBR5WG7vrQZuNqV4hP6BC_BY5sMpwx1v7j5IoU')
TRACKER_GID = int(os.getenv('TRACKER_GID', '2125924989'))
REVIEW_GID = int(os.getenv('REVIEW_GID', '1667806659'))
MAX_DETAILS = int(os.getenv('MAX_DETAILS_PER_RUN', '200'))
MAX_PAGES_PER_KEYWORD = int(os.getenv('MAX_PAGES_PER_KEYWORD', '10'))
DELAY = float(os.getenv('REQUEST_DELAY_SECONDS', '0.5'))

KEYWORDS = [
'Electronic Health Record','Electronic Medical Record','Health Information Exchange','Health Information Technology',
'Healthcare Information System','Clinical Information System','Hospital Information System','Electronic Visit Verification',
'Medicaid Management Information System','Public Health Information System','Healthcare Analytics','Healthcare Interoperability',
'Application Development','Application Modernization','Data Integration','API Integration','Cloud Migration',
'Managed Security Services','Healthcare Cybersecurity','IT Staff Augmentation','Artificial Intelligence','Machine Learning',
'Texas Health and Human Services Commission','Texas Department of State Health Services','Texas Medicaid','Integrated Eligibility'
]

TRACKER_HEADERS = [
'Opportunity ID','Date Found','State','Agency','Opportunity Title','Solicitation / Reference No.','Opportunity Type','Service Category',
'Official Source URL','Agency Site URL','Event ID','Event Type','Event Status','Purchase Type','Category Type','Government Type','Fiscal Year','Published Date','Deadline Date','Deadline Time / Time Zone','Days Until Deadline','Estimated Value (if published)','Scope Summary',
'Eligibility / Mandatory Requirements','Required Certifications / Experience','Contract Duration / Options','Incumbent / Prior Award (verified only)',
'Healthcare IT Fit (0-25)','Eligibility & Compliance (0-20)','Customer / Contract Fit (0-15)','Delivery Capability (0-15)',
'Deadline Feasibility (0-10)','Commercial Attractiveness (0-10)','Information Completeness (0-5)','Total Score (0-100)',
'Mandatory Gate (Pass / Fail / Unknown)','Recommendation (Go / Investigate / No-Go)','Fit Rationale','Risks / Gaps','Next Action',
'Last Verified Date','Owner','Review Status','Reviewer Notes','Search Keywords Matched','First Seen Date','Last Seen Date',
'Event Status','Event Start Date/Time','Buyer Contact Name','Buyer Contact Phone','Buyer Contact Email','Buyer Contact Address',
'Contact Extraction Status','Last Updated At','Record Hash'
]
REVIEW_HEADERS = ['Opportunity ID','Date Added to Review','State','Agency','Opportunity Title','Solicitation / Reference No.','Event Status',
'Official Source URL','Agency Site URL','Event ID','Event Type','Event Status','Purchase Type','Category Type','Government Type','Fiscal Year','Published Date','Deadline Date','Deadline Time / Time Zone','Days Until Deadline','Estimated Value (if published)','Scope Summary',
'Eligibility / Mandatory Requirements','Required Certifications / Experience','Healthcare IT Fit (0-25)','Eligibility & Compliance (0-20)',
'Customer / Contract Fit (0-15)','Delivery Capability (0-15)','Deadline Feasibility (0-10)','Commercial Attractiveness (0-10)',
'Information Completeness (0-5)','Total Score (0-100)','Mandatory Gate (Pass / Fail / Unknown)','Recommendation (Go / Investigate / No-Go)',
'Fit Rationale','Risks / Gaps','Next Action','Last Verified Date','Owner','Review Status','Reviewer Notes','Search Keywords Matched',
'Buyer Contact Name','Buyer Contact Phone','Buyer Contact Email','Buyer Contact Address','Contact Extraction Status']

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

def clean(s):
    return re.sub(r'\s+', ' ', (s or '')).strip()

def norm(s):
    return re.sub(r'[^a-z0-9]+', '', (s or '').lower())

def hash_record(row):
    fields = ['Opportunity Title','Agency','Event Status','Published Date','Deadline Date','Deadline Time / Time Zone','Scope Summary','Official Source URL','Buyer Contact Name','Buyer Contact Phone','Buyer Contact Email']
    return hashlib.sha256('|'.join(clean(row.get(k,'')) for k in fields).encode()).hexdigest()[:20]

def get_google_sheets():
    raw = os.getenv('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not raw:
        raise RuntimeError('Missing GOOGLE_SERVICE_ACCOUNT_JSON GitHub secret')
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive'])
    client = gspread.authorize(creds)
    book = client.open_by_key(SPREADSHEET_ID)
    tracker = next((w for w in book.worksheets() if w.id == TRACKER_GID), None)
    review = next((w for w in book.worksheets() if w.id == REVIEW_GID), None)
    if tracker is None: raise RuntimeError(f'Tracker worksheet gid {TRACKER_GID} not found')
    if review is None: raise RuntimeError(f'Manual Review worksheet gid {REVIEW_GID} not found')
    return tracker, review

def ensure_headers(ws, headers):
    """Ensure required headers exist without deleting/reordering existing columns.

    Automatically grows the Google worksheet grid before writing new headers.
    Existing column names and data are left untouched; missing names are appended
    to the next free columns. Safe to run on every scheduled execution.
    """
    from gspread.utils import rowcol_to_a1

    existing = ws.row_values(1)
    # Ignore only trailing blank header cells; preserve internal blank columns.
    while existing and not clean(existing[-1]):
        existing.pop()

    if not existing:
        required_cols = len(headers)
        if ws.col_count < required_cols:
            ws.add_cols(required_cols - ws.col_count)
        ws.update(range_name=f"A1:{rowcol_to_a1(1, required_cols)}",
                  values=[headers], value_input_option="RAW")
        return list(headers)

    # Treat existing names as authoritative and add only missing required names.
    # Also avoid adding a duplicate if the same header appears twice in `headers`.
    existing_names = {clean(h).casefold() for h in existing if clean(h)}
    missing = []
    for header in headers:
        if clean(header).casefold() not in existing_names:
            missing.append(header)
            existing_names.add(clean(header).casefold())

    if not missing:
        return existing

    start_col = len(existing) + 1
    end_col = start_col + len(missing) - 1
    if ws.col_count < end_col:
        # add_cols expands the grid without shrinking existing rows or columns.
        ws.add_cols(end_col - ws.col_count)

    start_cell = rowcol_to_a1(1, start_col)
    end_cell = rowcol_to_a1(1, end_col)
    ws.update(range_name=f"{start_cell}:{end_cell}", values=[missing], value_input_option="RAW")
    logging.info("Expanded worksheet '%s' as needed; added %d missing headers (%d total columns)",
                 ws.title, len(missing), end_col)
    return existing + missing

def rows_as_dicts(ws):
    vals = ws.get_all_values()
    if not vals: return [], []
    headers = vals[0]
    out=[]
    for r in vals[1:]:
        if not any(clean(x) for x in r): continue
        out.append({headers[i]: (r[i] if i < len(r) else '') for i in range(len(headers))})
    return headers, out

def upsert_sheet(ws, headers, records, key='Opportunity ID'):
    _, existing = rows_as_dicts(ws)
    index={clean(r.get(key,'')): i+2 for i,r in enumerate(existing) if clean(r.get(key,''))}
    updated=0; appended=0
    for rec in records:
        k=clean(rec.get(key,''))
        if not k: continue
        values=[rec.get(h,'') for h in headers]
        if k in index:
            
            from gspread.utils import rowcol_to_a1
            end_cell = rowcol_to_a1(index[k], len(headers))
            ws.update(range_name=f'A{index[k]}:{end_cell}', values=[values], value_input_option='USER_ENTERED'); updated += 1
        else:
            ws.append_row(values, value_input_option='USER_ENTERED'); appended += 1
    return appended, updated

def label_value_map(page):
    pairs={}
    # Capture common label/value tables and definition lists.
    for tr in page.locator('tr').all():
        try:
            cells=tr.locator('th,td').all_inner_texts()
            cells=[clean(x) for x in cells if clean(x)]
            if len(cells)>=2: pairs[cells[0].rstrip(':')]= ' | '.join(cells[1:])
        except Exception: pass
    for dt in page.locator('dt').all():
        try:
            key=clean(dt.inner_text()).rstrip(':'); val=clean(dt.locator('xpath=following-sibling::*[1]').inner_text())
            if key and val: pairs[key]=val
        except Exception: pass
    return pairs

def first_match(mapping, *terms):
    for k,v in mapping.items():
        if any(t in norm(k) for t in terms) and clean(v): return clean(v)
    return ''

def extract_event_id(url, mapping, text):
    for k,v in mapping.items():
        if 'eventid' in norm(k) or 'eventnumber' in norm(k):
            if v: return clean(v)
    m=re.search(r'Event ID\s+([A-Z0-9][A-Z0-9-]{4,})', text, re.I)
    if m:return m.group(1)
    q=parse_qs(urlparse(url).query)
    return clean((q.get('eSourceNumber') or q.get('eventId') or [''])[0])

def scrape_search_keyword(page, keyword):
    # Use the public registry landing page, not the old /index route.
    page.goto(SEARCH_URL, wait_until='domcontentloaded', timeout=60000)
    page.wait_for_timeout(1800)
    current_url = page.url
    body_preview = ''
    try:
        body_preview = clean(page.locator('body').inner_text(timeout=5000))[:1200]
    except Exception:
        pass
    if '/unsupported' in current_url.lower() or 'unsupported browser' in body_preview.lower():
        raise RuntimeError(f'GPR returned unsupported-browser page at {current_url}; page text: {body_preview[:500]}')
    # Match the visible field labelled Search Keyword; fall back to a text input.
    field = page.get_by_placeholder(re.compile('Enter Event ID or Title', re.I)).first
    if field.count()==0: field=page.locator('input[type="text"]').first
    field.fill(keyword)
    # Keep Event Status OPEN (as requested). Select by visible option where possible.
    for sel in page.locator('select').all():
        try:
            opts=[clean(x) for x in sel.locator('option').all_inner_texts()]
            if any(x.upper()=='OPEN' for x in opts): sel.select_option(label=re.compile('OPEN', re.I))
        except Exception: pass
    btn=page.get_by_role('button', name=re.compile('^Search$', re.I)).first
    if btn.count()==0: btn=page.locator('input[type="submit"]').first
    btn.click(timeout=15000)
    page.wait_for_load_state('domcontentloaded', timeout=30000)
    page.wait_for_timeout(800)
    found={}
    for page_no in range(MAX_PAGES_PER_KEYWORD):
        links=page.locator('a[href*="eventDetails" i], a[href*="eventdetails" i], a[href*="eSourceNumber" i]')
        for i in range(links.count()):
            try:
                href=links.nth(i).get_attribute('href')
                if href:
                    url=urljoin(page.url, href)
                    # Keep the original public URL format used by the registry; pur.doas links also work.
                    if '/gpr/eventDetails' in url and 'ssl.doas.state.ga.us' not in url:
                        url = url.replace('https://pur.doas.ga.gov/gpr/', BASE_URL)
                    # stable key: canonical event URL query, preserve actual official URL
                    key=parse_qs(urlparse(url).query).get('eSourceNumber',[url])[0]
                    found[key]={'url':url,'keyword':keyword}
            except Exception: pass
        # Try common next-page controls; stop when none is enabled.
        nxt=page.get_by_role('link', name=re.compile('^(Next|›|»)$', re.I)).last
        if nxt.count()==0:
            nxt=page.locator('a[aria-label*="Next"], button[aria-label*="Next"]').last
        if nxt.count()==0 or not nxt.is_visible(): break
        try:
            cls=(nxt.get_attribute('class') or '').lower(); aria=(nxt.get_attribute('aria-disabled') or '').lower()
            if 'disabled' in cls or aria=='true': break
            nxt.click(timeout=5000); page.wait_for_timeout(500)
        except Exception: break
    return found

def scrape_details(page, url, keywords):
    page.goto(url, wait_until='domcontentloaded', timeout=60000)
    page.wait_for_timeout(500)
    text=clean(page.locator('body').inner_text(timeout=15000))
    mapping=label_value_map(page)
    event_id=extract_event_id(url,mapping,text)
    # The visible page heading is the authoritative opportunity title (e.g. RFP828 - ...).
    title=''
    for sel in ['h1','h2','h3','.page-title','.event-title']:
        try:
            candidates=page.locator(sel).all_inner_texts()
            for candidate in candidates:
                candidate=clean(candidate)
                if re.match(r'^(RFP|RFQ|RFI|ITB|IFB|RFx)\s*[-#:]?\s*\w+', candidate, re.I):
                    title=candidate
                    break
            if title: break
        except Exception: pass
    if not title:
        # Registry page sometimes renders the title as a non-heading element.
        m=re.search(r'((?:RFP|RFQ|RFI|ITB|IFB)\s*\d+\s*[-:–—]\s*[^\n]{4,220}?)(?=\s+Start Date:|\s+End Date:|\s+Event Details)', text, re.I)
        if m: title=clean(m.group(1))
    if not title:
        title=first_match(mapping,'eventtitle','opportunitytitle','solicitationtitle','title')
        agency=first_match(mapping,'agency','governmententity','entityname','agencyname')
    if not agency:
        m=re.search(r'\b\d{5}\s+([A-Z][A-Z0-9 ,&().\-/]+?)(?:\s+Buyer Contact:|\s+Description)',text)
        if m: agency=clean(m.group(1))
    status=first_match(mapping,'eventstatus','status')
    if not status:
        m=re.search(r'Event Status\s+([A-Za-z ]+?)(?:\s+Purchase Type|\s+Category Type)',text,re.I)
        if m: status=clean(m.group(1))
    start=first_match(mapping,'startdate','eventstart')
    end=first_match(mapping,'enddate','closingdate','duedate','deadline')
    start_match=re.search(r'Start Date:\s*(.*?)(?:\s+End Date:|$)',text,re.I)
    end_match=re.search(r'End Date:\s*(.*?)(?:\s+Ends In:|\s+Event Details|$)',text,re.I)
    if start_match: start=clean(start_match.group(1))
    if end_match: end=clean(end_match.group(1))
    contact_name=first_match(mapping,'buyercontact','contactname','buyername')
    if not contact_name:
        m=re.search(r'Buyer Contact:\s*([^\d\n]+?)(?=\s+\d{3}[-.)]|\s+[\w.+-]+@[\w.-]+|$)',text,re.I)
        if m: contact_name=clean(m.group(1))
    email=first_match(mapping,'email','contactemail')
    if not email:
        m=re.search(r'[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}',text,re.I); email=m.group(0) if m else ''
    phone=first_match(mapping,'phone','telephone','contactphone')
    if not phone:
        m=re.search(r'(?<!\d)(?:\+?1[-. ]?)?\(?\d{3}\)?[-. ]\d{3}[-. ]\d{4}(?!\d)',text); phone=m.group(0) if m else ''
        if re.sub(r'\D','',phone) in {'5555555555','0000000000'}: phone=''
    desc=first_match(mapping,'description','eventdescription','scope')
    if not desc:
        m=re.search(r'Description\s+(.*?)(?:\s+NIGP Codes|\s+Documents|\s+Procurement Help Desk)',text,re.I)
        desc=clean(m.group(1)) if m else ''
    pub=first_match(mapping,'posteddate','publishdate','createddate','dateposted')
    value=first_match(mapping,'estimatedvalue','eventvalue','totalvalue')
    category=first_match(mapping,'categorytype','category','purchasetype','eventtype')
    event_type=first_match(mapping,'eventtype','type of response','response type')
    address=first_match(mapping,'address','buyeraddress','contactaddress')
    # The registry labels the external buyer's opportunity page as the Agency Site 'Link'.
    agency_site_url=''
    try:
        for a in page.locator('a[href]').all():
            label=clean(a.inner_text(timeout=500))
            href=a.get_attribute('href') or ''
            if label.lower() == 'link' and href and 'eventdetails' not in href.lower():
                agency_site_url=urljoin(page.url, href)
                break
    except Exception: pass
    # Explicitly extract the visible countdown shown beside the end date.
    days_until=''
    dm=re.search(r'Ends In:\s*(\d+)\s*days?', text, re.I)
    if dm: days_until=dm.group(1)
    event_id=event_id or ''
    return {
      'Opportunity ID':event_id or hashlib.sha1(url.encode()).hexdigest()[:12].upper(), 'Date Found':'', 'State':'Georgia',
      'Agency':agency,'Opportunity Title':title,'Solicitation / Reference No.':event_id,'Opportunity Type':event_type or category,
      'Service Category':category,'Official Source URL':url,'Agency Site URL':agency_site_url,'Event ID':event_id,'Published Date':pub,'Deadline Date':end,'Deadline Time / Time Zone':end,'Days Until Deadline':days_until,
      'Estimated Value (if published)':value,'Scope Summary':(desc[:3500] + ('\nEvent details: '+url+'#pane-eventDetails\nDocuments: '+url+'#pane-documents') if desc else 'Event details: '+url+'#pane-eventDetails\nDocuments: '+url+'#pane-documents'), 'Eligibility / Mandatory Requirements':'',
      'Required Certifications / Experience':'','Contract Duration / Options':'','Incumbent / Prior Award (verified only)':'',
      'Healthcare IT Fit (0-25)':'','Eligibility & Compliance (0-20)':'','Customer / Contract Fit (0-15)':'',
      'Delivery Capability (0-15)':'','Deadline Feasibility (0-10)':'','Commercial Attractiveness (0-10)':'',
      'Information Completeness (0-5)':'','Total Score (0-100)':'','Mandatory Gate (Pass / Fail / Unknown)':'Unknown',
      'Recommendation (Go / Investigate / No-Go)':'Investigate','Fit Rationale':'Automated keyword discovery; human review required.',
      'Risks / Gaps':'Confirm healthcare/IT relevance, requirements, deadline and eligibility in the official event documents.',
      'Next Action':'Open official event; review documents and complete scoring.','Last Verified Date':datetime.now().strftime('%Y-%m-%d'),
      'Owner':'','Review Status':'Needs Review','Reviewer Notes':'','Search Keywords Matched':'; '.join(sorted(set(keywords))),
      'First Seen Date':'','Last Seen Date':'','Event Status':status,'Event Start Date/Time':start,'Buyer Contact Name':contact_name,
      'Buyer Contact Phone':phone,'Buyer Contact Email':email,'Buyer Contact Address':address,
      'Contact Extraction Status':'Contact found' if (contact_name or phone or email) else 'Not found on page',
      'Last Updated At':datetime.now().isoformat(timespec='seconds'),'Record Hash':''}

def is_complete_record(row):
    """Return true only when the core solicitation fields are present and internally usable."""
    title = clean(row.get('Opportunity Title', ''))
    agency = clean(row.get('Agency', ''))
    deadline = clean(row.get('Deadline Date', ''))
    scope = clean(row.get('Scope Summary', ''))
    url = clean(row.get('Official Source URL', ''))
    invalid_title = not title or title.lower() in {'event', 'announcement', 'untitled'}
    invalid_agency = not agency or agency.lower() in {'unknown', 'n/a', 'none'}
    has_scope = bool(scope and len(re.sub(r'https?://\S+', '', scope).strip()) >= 30)
    has_deadline = bool(deadline and re.search(r'20\d{2}|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec', deadline, re.I))
    valid_url = 'eventdetails' in url.lower() and 'esourcenumber=' in url.lower()
    return (not invalid_title and not invalid_agency and (has_scope or has_deadline) and valid_url)

def export_data(records, review_records, diagnostics):
    os.makedirs('data', exist_ok=True)
    for basename, rows in [('georgia_leads', records), ('georgia_manual_review', review_records)]:
        with open(f'data/{basename}.json', 'w', encoding='utf-8') as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        if rows:
            columns = list(dict.fromkeys(k for row in rows for k in row.keys()))
        else:
            columns = TRACKER_HEADERS if basename == 'georgia_leads' else REVIEW_HEADERS
        with open(f'data/{basename}.csv', 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=columns, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(rows)
    with open('data/georgia_diagnostics.json', 'w', encoding='utf-8') as f:
        json.dump(diagnostics, f, ensure_ascii=False, indent=2)

def main():
    tracker, review=get_google_sheets()
    tracker_headers=ensure_headers(tracker,TRACKER_HEADERS)
    review_headers=ensure_headers(review,REVIEW_HEADERS)
    _, old_tracker=rows_as_dicts(tracker)
    old_by_id={clean(r.get('Opportunity ID','')):r for r in old_tracker if clean(r.get('Opportunity ID',''))}
    found={}; failures=[]
    with sync_playwright() as p:
        browser=p.chromium.launch(headless=True)
        context=browser.new_context(
            viewport={'width':1440,'height':1000},
            locale='en-US',
            timezone_id='America/New_York',
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
            extra_http_headers={'Accept-Language':'en-US,en;q=0.9'},
        )
        page=context.new_page()
        for i,keyword in enumerate(KEYWORDS,1):
            try:
                results=scrape_search_keyword(page,keyword)
                logging.info('[%s/%s] %s => %s results',i,len(KEYWORDS),keyword,len(results))
                for k,v in results.items():
                    if k in found: found[k]['keywords'].add(keyword)
                    else: found[k]={'url':v['url'],'keywords':{keyword}}
            except Exception as e:
                logging.exception('Search failed for keyword: %s',keyword); failures.append({'keyword':keyword,'error':str(e)})
            time.sleep(DELAY)
        # Revisit newly discovered records and existing tracker records daily, so later edits are detected.
        urls=[]
        for key,item in found.items(): urls.append((key,item))
        known_urls = {clean(item['url']) for _, item in urls}
        for old in old_tracker:
            old_url = clean(old.get('Official Source URL',''))
            old_id = clean(old.get('Opportunity ID',''))
            if old_url and old_id and old_url not in known_urls:
                urls.append((old_id, {'url': old_url, 'keywords': set(filter(None, clean(old.get('Search Keywords Matched','')).split(';')))}))
                known_urls.add(old_url)
        if len(urls)>MAX_DETAILS:
            # Prioritize new records, then records with oldest verification date.
            urls.sort(key=lambda kv: (0 if kv[0] not in old_by_id else 1, old_by_id.get(kv[0],{}).get('Last Verified Date','')))
            urls=urls[:MAX_DETAILS]
        records=[]
        for n,(key,item) in enumerate(urls,1):
            try:
                row=scrape_details(page,item['url'],item['keywords'])
                oid=clean(row['Opportunity ID'])
                prior=old_by_id.get(oid,{})
                if not prior:
                    prior=next((old for old in old_tracker if clean(old.get('Official Source URL','')).split('#')[0] == clean(item['url']).split('#')[0]), {})
                row['Date Found']=prior.get('Date Found') or datetime.now().strftime('%Y-%m-%d')
                row['First Seen Date']=prior.get('First Seen Date') or row['Date Found']
                row['Last Seen Date']=datetime.now().strftime('%Y-%m-%d')
                # Preserve existing verified values when the page extraction is blank, plus human-entered review fields.
                for col in ['Healthcare IT Fit (0-25)','Eligibility & Compliance (0-20)','Customer / Contract Fit (0-15)','Delivery Capability (0-15)',
                    'Deadline Feasibility (0-10)','Commercial Attractiveness (0-10)','Information Completeness (0-5)','Total Score (0-100)',
                    'Mandatory Gate (Pass / Fail / Unknown)','Recommendation (Go / Investigate / No-Go)','Fit Rationale','Risks / Gaps','Next Action','Owner','Review Status','Reviewer Notes',
                    'Eligibility / Mandatory Requirements','Required Certifications / Experience','Contract Duration / Options','Incumbent / Prior Award (verified only)']:
                    if clean(prior.get(col,'')) and not clean(row.get(col,'')): row[col]=prior[col]
                for col in ['Healthcare IT Fit (0-25)','Eligibility & Compliance (0-20)','Customer / Contract Fit (0-15)','Delivery Capability (0-15)',
                    'Deadline Feasibility (0-10)','Commercial Attractiveness (0-10)','Information Completeness (0-5)','Total Score (0-100)',
                    'Mandatory Gate (Pass / Fail / Unknown)','Recommendation (Go / Investigate / No-Go)','Fit Rationale','Risks / Gaps','Next Action','Owner','Review Status','Reviewer Notes',
                    'Eligibility / Mandatory Requirements','Required Certifications / Experience','Contract Duration / Options','Incumbent / Prior Award (verified only)']:
                    if clean(prior.get(col,'')): row[col]=prior[col]
                row['Record Hash']=hash_record(row)
                records.append(row)
                logging.info('Details %s/%s: %s',n,len(urls),oid)
            except Exception as e:
                logging.exception('Detail failed for %s',item['url']); failures.append({'url':item['url'],'error':str(e)})
            time.sleep(DELAY)
        browser.close()
    lead_records=[]
    review_records=[]
    for r in records:
        if is_complete_record(r):
            lead_records.append(r)
        else:
            review_records.append(r)
    # Tracker contains only records passing core-field validation; ambiguous/incomplete records go to Manual Review.
    if lead_records:
        added,updated=upsert_sheet(tracker,tracker_headers,lead_records)
        logging.info('Tracker sync: %s added, %s updated',added,updated)
    else:
        logging.info('Tracker sync: no records passed completeness checks; records will be written to Manual Review')
    if review_records:
        mapped=[{h:r.get(h,'') for h in review_headers} for r in review_records]
        ra,ru=upsert_sheet(review,review_headers,mapped)
        logging.info('Manual review sync: %s added, %s updated',ra,ru)
    else:
        logging.info('Manual review sync: no records to add/update')
    diagnostics={'run_at':datetime.now(timezone.utc).isoformat(),'keywords_total':len(KEYWORDS),'unique_events_discovered':len(found),
      'details_attempted':len(urls) if 'urls' in locals() else 0,'records_scraped':len(records),'tracker_records':len(lead_records),
      'manual_review_records':len(review_records),'search_failures':failures,'tracker_gid':TRACKER_GID,'review_gid':REVIEW_GID,'max_details_per_run':MAX_DETAILS}
    _, tracker_snapshot = rows_as_dicts(tracker)
    _, review_snapshot = rows_as_dicts(review)
    export_data(tracker_snapshot, review_snapshot, diagnostics)
    for filename in ['data/georgia_leads.csv','data/georgia_leads.json','data/georgia_manual_review.csv','data/georgia_manual_review.json','data/georgia_diagnostics.json']:
        logging.info('Data export ready: %s (%s bytes)', filename, os.path.getsize(filename))
    logging.info('Diagnostics: %s',json.dumps(diagnostics))
    if not found and not records:
        raise RuntimeError('No event-detail links were found and no existing events could be refreshed. Inspect data/gpr_debug.txt and data/gpr_debug.html from the workflow artifact.')

if __name__=='__main__': main()
