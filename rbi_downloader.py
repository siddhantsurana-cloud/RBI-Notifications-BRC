"""
RBI Notifications / Guidelines Downloader + Per-Circular Email Summaries
-----------------------------------------------------------------------
For every NEW RBI Notification, Master Direction, Master Circular and
Press Release it:
  1. downloads the PDF into  S:\\RBI - Notifications\\FY yyyy-yy\\yyyy-mm-dd\\<Category>\\
  2. writes a brief summary (key points, applicability, effective date,
     action required) - using Claude AI if an API key is set, otherwise
     a simple extract of the opening paragraphs
  3. emails ONE mail per circular, with the PDF attached, to EMAIL_TO

Install once:
    pip install requests beautifulsoup4 feedparser truststore pypdf
"""

import html
import json
import logging
import os
import re
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import urljoin

import feedparser
import requests
from bs4 import BeautifulSoup
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:                                   # Windows certificate store (office networks)
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

# ============================ CONFIG ============================
BASE_DIR = Path(r"S:\RBI - Notifications")
# If S: is a mapped network drive and the task runs while logged off, use the
# UNC path instead, e.g.  Path(r"\\fileserver\share\RBI - Notifications")

FEEDS = {
    "Notifications":   "https://www.rbi.org.in/notifications_rss.xml",
    "Press_Releases":  "https://www.rbi.org.in/pressreleases_rss.xml",
}
LISTING_PAGES = {
    "Notifications":      "https://www.rbi.org.in/Scripts/NotificationUser.aspx",
    "Master_Directions":  "https://www.rbi.org.in/Scripts/BS_ViewMasterDirections.aspx",
    "Master_Circulars":   "https://www.rbi.org.in/Scripts/BS_ViewMasterCirculardetails.aspx",
}
DELAY_SECONDS = 5                        # pause between requests (be gentle with RBI)
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
}
MAX_EMAILS_PER_RUN = 25                  # more than this -> one digest email instead

# ---- Email (one mail per circular) ----
EMAIL_ENABLED = False                    # emails paused; set True to resume
EMAIL_FROM    = "Siddhant@bihanico.in"
EMAIL_TO      = ["Siddhant@bihanico.in", "Rashmi@bihanico.in"]
# Settings taken from Outlook (bihanico.in mail server)
SMTP_SERVER   = "mail.bihanico.in"
SMTP_PORT     = 587
SMTP_USER     = "siddhant@bihanico.in"   # same login as in Outlook
SMTP_PASSWORD = os.environ.get("RBI_SMTP_PASSWORD", "PUT-PASSWORD-HERE")
MAX_ATTACH_MB = 20                       # bigger PDFs are linked, not attached

# ---- Summaries ----
# Paste a Claude API key (console.anthropic.com) for proper AI summaries.
# Leave empty to use a simple extract of the circular's opening text.
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL      = "claude-sonnet-5-5"
# ================================================================

STATE_FILE = BASE_DIR / "downloaded.json"
BASELINE_FILE = BASE_DIR / "baseline_complete.txt"
INDEX_FILE = BASE_DIR / "index.json"          # full titles + RBI links for the dashboard
DASHBOARD_FILE = BASE_DIR / "RBI Dashboard.html"   # emails start only after this exists
try:
    BASE_DIR.mkdir(parents=True, exist_ok=True)
except Exception as e:
    print(f"ERROR: cannot reach the folder {BASE_DIR}\n  {e}")
    raise SystemExit(1)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(BASE_DIR / "rbi_downloader.log", encoding="utf-8"),
              logging.StreamHandler()],
)
logging.getLogger("pypdf").setLevel(logging.ERROR)      # hide noisy PDF-parser messages
session = requests.Session()
session.headers.update(HEADERS)
_retry = Retry(total=2, backoff_factor=3, status_forcelist=[429, 500, 502, 503, 504])
session.mount("https://", HTTPAdapter(max_retries=_retry))
session.mount("http://", HTTPAdapter(max_retries=_retry))
NEW_ITEMS = []
FAILS = {"in_a_row": 0}


class RBIBlocking(Exception):
    """RBI's server keeps refusing connections - stop and resume next run."""


def note_success():
    FAILS["in_a_row"] = 0


def note_failure():
    FAILS["in_a_row"] += 1
    if FAILS["in_a_row"] == 4:
        logging.warning("RBI server refusing connections - pausing 3 minutes...")
        time.sleep(180)
    elif FAILS["in_a_row"] >= 8:
        raise RBIBlocking()

MONTHS = r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[a-z]*\.?"
DATE_PATTERNS = [
    (re.compile(MONTHS + r"\s+\d{1,2},\s*\d{4}"), ["%B %d, %Y", "%b %d, %Y", "%b. %d, %Y"]),
    (re.compile(r"\d{1,2}\s+" + MONTHS + r",?\s+\d{4}"), ["%d %B %Y", "%d %b %Y", "%d %B, %Y", "%d %b, %Y"]),
]
REF_PATTERN = re.compile(r"RBI/\d{4}-\d{2,4}/\d+[^\n]{0,80}", re.I)

# PDFs linked on every RBI page (site menus/footer), not part of any circular
SITE_WIDE_PDF = re.compile(r"(utkarsh|accessibility)", re.I)
GLOBAL_PDFS = set()          # filled from the RBI home page at start-up


# ---------------------------------------------------------------- helpers
def load_state():
    return set(json.loads(STATE_FILE.read_text())) if STATE_FILE.exists() else set()


def save_state(done):
    tmp = STATE_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(sorted(done), indent=1))
    tmp.replace(STATE_FILE)


def fix_text(text):
    """Repair garbled characters such as 'â€“' (should be an en dash)."""
    if any(m in text for m in ("â€", "Ã", "Â")):
        try:
            return text.encode("cp1252").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
    return text


def safe_name(text, limit=100):
    text = fix_text(text)
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", text)
    text = re.sub(r"\s+", " ", text).strip().rstrip(".")
    return text[:limit] or "untitled"


def financial_year(d):
    start = d.year if d.month >= 4 else d.year - 1
    return f"FY {start}-{str(start + 1)[-2:]}"


def parse_date(text):
    for pattern, formats in DATE_PATTERNS:
        for m in pattern.finditer(text):
            s = m.group(0).replace("Sept", "Sep")
            for fmt in formats:
                try:
                    return datetime.strptime(s, fmt).date()
                except ValueError:
                    continue
    return None


def inspect_page(page_url):
    """Return (pdf_links, publication_date, page_text) for an RBI detail page."""
    r = session.get(page_url, timeout=60)
    r.raise_for_status()
    soup = BeautifulSoup(r.content, "html.parser")
    links = [urljoin(page_url, a["href"]) for a in soup.find_all("a", href=True)
             if a["href"].lower().split("?")[0].endswith(".pdf")]
    links = [l for l in links
             if l not in GLOBAL_PDFS and not SITE_WIDE_PDF.search(l.rsplit("/", 1)[-1])]
    for t in soup(["script", "style", "nav", "header", "footer"]):
        t.decompose()
    main = soup.find(class_=re.compile("tablebg|content", re.I)) or soup
    text = main.get_text("\n", strip=True)
    pub = parse_date(text) or parse_date(soup.get_text(" ", strip=True))
    return list(dict.fromkeys(links)), pub, text


def pdf_text(path, max_pages=15):
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        return "\n".join((p.extract_text() or "") for p in reader.pages[:max_pages])
    except Exception as e:
        logging.warning("Could not read PDF text %s: %s", path, e)
        return ""


def is_real_pdf(path):
    try:
        with open(path, "rb") as f:
            return f.read(1024).lstrip().startswith(b"%PDF")
    except OSError:
        return False


def fetch_pdf_bytes(url, referer):
    """Download a PDF; returns bytes only if the server really sent a PDF."""
    hdrs = {"Referer": referer, "Accept": "application/pdf,*/*"}
    for attempt in range(2):
        r = session.get(url, headers=hdrs, timeout=120)
        r.raise_for_status()
        if r.content.lstrip()[:4] == b"%PDF":
            return r.content
        # RBI sometimes returns an HTML page; it may contain the real PDF link
        soup = BeautifulSoup(r.content, "html.parser")
        alt = [urljoin(r.url, a["href"]) for a in soup.find_all("a", href=True)
               if a["href"].lower().split("?")[0].endswith(".pdf")
               and not SITE_WIDE_PDF.search(a["href"])]
        meta = soup.find("meta", attrs={"http-equiv": re.compile("refresh", re.I)})
        if meta and "url=" in meta.get("content", "").lower():
            alt.insert(0, urljoin(r.url, meta["content"].split("=", 1)[1].strip("'\" ")))
        if alt and alt[0] != url:
            url = alt[0]
        else:
            time.sleep(5 * (attempt + 1))
    title = BeautifulSoup(r.content, "html.parser").title
    logging.warning("Not a PDF (server sent a web page '%s'): %s",
                    title.get_text(strip=True)[:80] if title else "", url)
    return None


def download(url, folder, title, referer=None):
    folder.mkdir(parents=True, exist_ok=True)
    fname = safe_name(title) + "__" + url.rsplit("/", 1)[-1].split("?")[0]
    if not fname.lower().endswith(".pdf"):
        fname += ".pdf"
    dest = folder / fname
    if dest.exists() and is_real_pdf(dest):
        return dest
    data = fetch_pdf_bytes(url, referer or "https://www.rbi.org.in/")
    if data is None:
        if dest.exists():
            dest.unlink()                    # remove an old broken copy
        return None
    dest.write_bytes(data)
    logging.info("Saved %s", dest)
    return dest


# ---------------------------------------------------------------- summaries
SUMMARY_PROMPT = """You are a compliance analyst at an Indian regulated entity.
Summarise the RBI document below for busy senior staff. Respond ONLY with JSON
(no markdown fences) using exactly these keys:
{{"summary": "2-3 sentence overview in plain English",
 "key_points": ["5-8 short bullet points with the substantive changes / requirements"],
 "applicable_to": "which entities it applies to",
 "effective_date": "effective date or 'With immediate effect' or 'Not specified'",
 "action_required": "what a regulated entity needs to do, or 'For information only'"}}

Title: {title}
Category: {category}

Document text:
{text}"""


def summarise_with_claude(title, category, text):
    r = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={"x-api-key": ANTHROPIC_API_KEY,
                 "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
        json={"model": CLAUDE_MODEL, "max_tokens": 1500,
              "messages": [{"role": "user", "content": SUMMARY_PROMPT.format(
                  title=title, category=category, text=text[:60000])}]},
        timeout=180,
    )
    r.raise_for_status()
    out = "".join(b.get("text", "") for b in r.json().get("content", []))
    out = re.sub(r"^```(json)?|```$", "", out.strip()).strip()
    return json.loads(out)


def simple_summary(text):
    """No-AI fallback: first meaningful paragraphs of the circular."""
    paras = [p.strip() for p in text.split("\n")
             if len(p.strip()) > 80 and not p.strip().lower().startswith(("dear", "madam", "yours"))]
    return {"summary": paras[0][:600] if paras else "Please refer to the attached PDF.",
            "key_points": [p[:350] for p in paras[1:7]],
            "applicable_to": "", "effective_date": "", "action_required": ""}


def make_summary(item):
    text = item["text"]
    if len(text) < 400:                       # page text too thin -> read the PDF
        for f in item["files"]:
            text += "\n" + pdf_text(f)
    if ANTHROPIC_API_KEY:
        try:
            return summarise_with_claude(item["title"], item["category"], text)
        except Exception as e:
            logging.error("AI summary failed for %s: %s", item["title"], e)
    return simple_summary(text)


# ---------------------------------------------------------------- core
def process_item(category, title, page_url, done, hint_date=None):
    if not page_url or page_url in done:
        return
    title = fix_text(title)
    try:
        if page_url.lower().split("?")[0].endswith(".pdf"):
            pdfs, page_date, text = [page_url], None, ""
        else:
            pdfs, page_date, text = inspect_page(page_url)
        pub_date = hint_date or page_date or date.today()
        folder = BASE_DIR / financial_year(pub_date) / pub_date.isoformat() / category
        files = []
        for pdf in pdfs:
            f = download(pdf, folder, title, referer=page_url)
            if f:
                files.append(f)
            time.sleep(DELAY_SECONDS)
        if pdfs and not files:
            logging.error("No valid PDF obtained for %s - will retry next run", title)
            note_failure()
            return                               # not marked done -> retried later
        note_success()
        ref = REF_PATTERN.search(text)
        done.add(page_url)
        save_state(done)
        record_index(pub_date, category, title, page_url, ref.group(0).strip() if ref else "")                         # save progress after every item
        NEW_ITEMS.append({"category": category, "title": title, "url": page_url,
                          "date": pub_date, "files": files, "text": text,
                          "ref": ref.group(0).strip() if ref else ""})
    except RBIBlocking:
        raise
    except Exception as e:
        logging.error("Failed %s (%s): %s", title, page_url, e)
        note_failure()


def record_index(pub_date, category, title, url, ref):
    try:
        idx = json.loads(INDEX_FILE.read_text(encoding="utf-8")) if INDEX_FILE.exists() else {}
    except Exception:
        idx = {}
    idx[f"{pub_date.isoformat()}|{category}|{safe_name(title)}"] = {"t": title, "u": url, "r": ref[:120]}
    INDEX_FILE.write_text(json.dumps(idx, ensure_ascii=False, indent=0), encoding="utf-8")


def from_rss(done):
    for category, feed_url in FEEDS.items():
        try:
            r = session.get(feed_url, timeout=60)
            r.raise_for_status()
            feed = feedparser.parse(r.content)
        except Exception as e:
            logging.error("Could not read feed %s: %s", feed_url, e)
            continue
        for e in feed.entries:
            hint = date(*e.published_parsed[:3]) if e.get("published_parsed") else None
            process_item(category, e.get("title", "untitled"), e.get("link", ""), done, hint)


def from_listing_pages(done):
    for category, url in LISTING_PAGES.items():
        try:
            r = session.get(url, timeout=60)
            r.raise_for_status()
        except Exception as e:
            logging.error("Listing page failed %s: %s", url, e)
            continue
        soup = BeautifulSoup(r.content, "html.parser")
        for a in soup.find_all("a", href=True):
            href = urljoin(url, a["href"])
            title = a.get_text(" ", strip=True)
            if title and re.search(r"(Id=\d+|\.pdf$)", href, re.I):
                row = a.find_parent("tr")
                hint = parse_date(row.get_text(" ", strip=True)) if row else None
                process_item(category, title, href, done, hint)


# ---------------------------------------------------------------- email
def build_email(item, s):
    from email.message import EmailMessage
    cat = item["category"].replace("_", " ")
    e = html.escape
    msg = EmailMessage()
    msg["Subject"] = f"RBI {cat} | {item['date']:%d-%b-%Y} | {item['title'][:150]}"
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(EMAIL_TO)

    def row(label, value):
        return (f"<tr><td style='padding:4px 10px;color:#555'><b>{label}</b></td>"
                f"<td style='padding:4px 10px'>{value}</td></tr>") if value else ""

    points = "".join(f"<li style='margin-bottom:4px'>{e(p)}</li>" for p in s.get("key_points", []))
    links_only = [f for f in item["files"] if f and f.stat().st_size > MAX_ATTACH_MB * 1e6]
    html_body = f"""
<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#222">
 <h3 style="margin:0 0 8px;color:#0b3d6e">{e(item['title'])}</h3>
 <table style="border-collapse:collapse;margin-bottom:12px">
  {row('Category', e(cat))}
  {row('Date', f"{item['date']:%d %B %Y}")}
  {row('Financial Year', financial_year(item['date']))}
  {row('Reference', e(item['ref']))}
  {row('Applicable to', e(s.get('applicable_to', '')))}
  {row('Effective date', e(s.get('effective_date', '')))}
 </table>
 <p><b>Summary</b><br>{e(s.get('summary', ''))}</p>
 {"<p><b>Key points</b></p><ul>" + points + "</ul>" if points else ""}
 {"<p><b>Action required</b><br>" + e(s.get('action_required', '')) + "</p>" if s.get('action_required') else ""}
 <p><a href="{e(item['url'])}">View on RBI website</a></p>
 {"<p><i>PDF too large to attach; please use the link above.</i></p>" if links_only else ""}
 <p style="color:#888;font-size:12px">Automated summary - please refer to the attached
 RBI document for the authoritative text. Saved under {e(str(BASE_DIR))}.</p>
</div>"""
    plain = (f"{item['title']}\n{cat} | {item['date']:%d-%b-%Y} | {financial_year(item['date'])}\n\n"
             f"Summary: {s.get('summary', '')}\n\n" +
             "\n".join(f"- {p}" for p in s.get("key_points", [])) +
             f"\n\nAction required: {s.get('action_required', '')}\n{item['url']}\n")
    msg.set_content(plain)
    msg.add_alternative(html_body, subtype="html")
    for f in item["files"]:
        if f and f.exists() and f not in links_only and is_real_pdf(f):
            msg.add_attachment(f.read_bytes(), maintype="application",
                               subtype="pdf", filename=f.name.split("__")[-1])
    return msg


def send_digest(smtp, items):
    from email.message import EmailMessage
    msg = EmailMessage()
    msg["Subject"] = f"RBI Update {date.today():%d-%b-%Y}: {len(items)} new items (digest)"
    msg["From"] = EMAIL_FROM
    msg["To"] = ", ".join(EMAIL_TO)
    rows = "".join(
        f"<tr><td>{i['date']:%d-%b-%Y}</td><td>{html.escape(i['category'].replace('_', ' '))}</td>"
        f"<td><a href='{html.escape(i['url'])}'>{html.escape(i['title'])}</a></td></tr>"
        for i in sorted(items, key=lambda i: i["date"], reverse=True))
    msg.set_content("\n".join(f"{i['date']:%d-%b-%Y} {i['title']} {i['url']}" for i in items))
    msg.add_alternative(f"<p>{len(items)} new RBI items (too many for individual mails). "
                        f"PDFs are saved under {html.escape(str(BASE_DIR))}.</p>"
                        "<table border='1' cellpadding='6' cellspacing='0'>"
                        f"<tr><th>Date</th><th>Category</th><th>Title</th></tr>{rows}</table>",
                        subtype="html")
    smtp.send_message(msg)
    logging.info("Digest email sent (%d items)", len(items))


def send_emails(items):
    import smtplib
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60) as smtp:
        smtp.ehlo()
        if smtp.has_extn("starttls"):          # encrypt if the server supports it
            try:
                import ssl
                ctx = ssl.create_default_context()
                ctx.check_hostname = False     # many hosted mail servers use a
                ctx.verify_mode = ssl.CERT_NONE  # certificate for another hostname
                smtp.starttls(context=ctx)
                smtp.ehlo()
            except Exception as e:
                logging.warning("STARTTLS failed, continuing without encryption: %s", e)
                smtp.close()
                smtp.connect(SMTP_SERVER, SMTP_PORT)
                smtp.ehlo()
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        if len(items) > MAX_EMAILS_PER_RUN:
            send_digest(smtp, items)
            return
        for item in sorted(items, key=lambda i: i["date"]):
            try:
                summary = make_summary(item)
                smtp.send_message(build_email(item, summary))
                logging.info("Emailed: %s", item["title"])
                time.sleep(3)
            except Exception as e:
                logging.error("Email failed for %s: %s", item["title"], e)


# ---------------------------------------------------------------- classification
# Subject groups: (name, keywords as a regular expression). Edit freely.
TOPIC_RULES = [
    ("FDI",                         r"foreign direct investment|\bFDI\b|non-debt instrument|downstream investment|foreign investment in india"),
    ("ODI",                         r"overseas direct investment|overseas investment|\bODI\b|\bOPI\b|overseas portfolio investment"),
    ("ECB & Trade Credit",          r"external commercial borrowing|\bECBs?\b|trade credit"),
    ("LRS & Remittances",           r"liberali[sz]ed remittance|\bLRS\b|inward remittance|outward remittance|money transfer service"),
    ("Exports & Imports",           r"\bexports?\b|\bimports?\b|EDPMS|IDPMS|merchanting trade|SOFTEX"),
    ("Forex Markets & Derivatives", r"foreign exchange market|forex|derivative|hedging|OTC|currency futures|market-makers?|interest rate futures"),
    ("FEMA - General",              r"\bFEMA\b|foreign exchange management|compounding|authori[sz]ed (dealer|person)|A\.\s?P\.\s?\(DIR"),
    ("NBFC",                        r"non-banking financial|\bNBFCs?\b|housing finance compan|\bHFCs?\b|asset reconstruction|core investment compan|microfinance"),
    ("Co-operative Banks",          r"co-?operative banks?|\bUCBs?\b|\bStCBs?\b|\bDCCBs?\b"),
    ("Small Finance & Payments Banks", r"small finance banks?|payments banks?|local area banks?"),
    ("KYC & AML",                   r"\bKYC\b|know your customer|money laundering|\bAML\b|\bCFT\b|UAPA|unlawful activities|sanctions list|\bPMLA\b"),
    ("Capital & Prudential Norms",  r"capital adequacy|basel|prudential norms?|liquidity coverage|\bLCR\b|\bNSFR\b|leverage ratio|provisioning|asset classification|income recognition"),
    ("Loans & Credit",              r"loans and advances|lending|credit facilit|gold loan|co-lending|education loan|housing loan|\bEBLR\b|\bMCLR\b|large exposure|credit risk|credit information"),
    ("Deposits & Interest Rates",   r"\bdeposits?\b|interest rates? on|savings account|\bFCNR|\bNRE\b|\bNRO\b"),
    ("Priority Sector & Inclusion", r"priority sector|financial inclusion|lead bank|\bMSMEs?\b|micro, small|agricultur|kisan|scheduled castes|scheduled tribes|self help group|\bSHGs?\b|PMJDY|NRLM"),
    ("Payment Systems",             r"payment systems?|\bUPI\b|\bNEFT\b|\bRTGS\b|prepaid payment|\bPPIs?\b|card networks?|debit cards?|credit cards?|payment aggregator|Bharat Bill|\bNACH\b|cheque truncation"),
    ("G-Secs & Money Market",       r"government securit|\bG-?secs?\b|treasury bills?|money market|\brepo\b|reverse repo|state development loans?|\bSDLs?\b|call money|certificates? of deposit|commercial paper|liquidity adjustment|\bLAF\b|\bSDF\b|\bMSF\b|open market operation|\bOMOs?\b|variable rate|auction"),
    ("Monetary Policy",             r"monetary policy|policy repo rate|\bMPC\b|bank rate"),
    ("Currency & Coins",            r"banknotes?|\bcoins?\b|legal tender|soiled notes?|currency chest|currency management"),
    ("Customer Service & Grievance", r"customer service|grievance|ombudsman|consumer protection|\bfrauds?\b|unclaimed deposit|nomination|complaints?"),
    ("Penalties & Enforcement",     r"monetary penalty|imposes penalty|penalty on|cancels? (the )?certificate of registration|cancell?ation of (licence|registration)|section 35A|all inclusive directions|restrictions on"),
    ("Governance & Audit",          r"corporate governance|statutory (central )?audit|auditors?|whole[- ]time director|fit and proper|internal audit|concurrent audit|compensation"),
    ("IT, Cyber & Outsourcing",     r"cyber|information technology|\bIT\b governance|outsourcing|cloud"),
    ("Commercial Banks",            r"commercial banks?|scheduled banks?"),
]
TOPIC_RULES = [(n, re.compile(rx, re.I)) for n, rx in TOPIC_RULES]
FOREX_TOPICS = {"FDI", "ODI", "ECB & Trade Credit", "LRS & Remittances", "Exports & Imports",
                "Forex Markets & Derivatives", "FEMA - General"}

DEPT_RULES = [
    ("FED - Foreign Exchange",          r"A\.\s?P\.\s?\(DIR Series\)|\bFED\b|Foreign Exchange Department"),
    ("DOR - Regulation",                r"\bD[Oo]R\b|Department of Regulation|\bDBR\b|\bDNBR\b"),
    ("DoS - Supervision",               r"\bD[Oo]S\b|Department of Supervision"),
    ("DPSS - Payment Systems",          r"\bDPSS\b|Payment and Settlement Systems"),
    ("FIDD - Inclusion & Development",  r"\bFIDD\b|Financial Inclusion and Development"),
    ("FMRD - Financial Markets Reg.",   r"\bFMRD\b|Financial Markets Regulation"),
    ("FMOD - Financial Markets Ops.",   r"\bFMOD\b|Financial Markets Operations"),
    ("IDMD - Internal Debt Mgmt.",      r"\bIDMD\b|Internal Debt Management"),
    ("DCM - Currency Management",       r"\bDCM\b|Department of Currency Management"),
    ("CEPD - Consumer Protection",      r"\bCEPD\b|Consumer Education and Protection"),
    ("DGBA - Govt & Bank Accounts",     r"\bDGBA\b|Government and Bank Accounts"),
    ("MPD - Monetary Policy",           r"\bMPD\b|Monetary Policy Department"),
    ("DoE - Enforcement",               r"\bEnforcement Department\b|\bEFD\b"),
]
DEPT_RULES = [(n, re.compile(rx)) for n, rx in DEPT_RULES]
CLASS_CACHE = BASE_DIR / "classification.json"
CLASS_VERSION = 1


def classify(title, text, category):
    """Return (topics, department, reference) for one publication."""
    head = text[:3000]
    topics = []
    for name, rx in TOPIC_RULES:                    # title matches first
        if rx.search(title):
            topics.append(name)
    scored = []
    for name, rx in TOPIC_RULES:                    # then strong matches in the text
        if name not in topics:
            n = len(rx.findall(head))
            if n >= 2:
                scored.append((n, name))
    topics += [n for _, n in sorted(scored, reverse=True)]
    # a specific forex topic makes the generic one redundant
    if "FEMA - General" in topics and len(FOREX_TOPICS & set(topics)) > 1:
        topics.remove("FEMA - General")
    # the generic entity tag only if nothing more specific was found
    if "Commercial Banks" in topics and len(topics) > 1:
        topics.remove("Commercial Banks")
    topics = topics[:3] or ["Others"]

    dept, best = "", 10 ** 9
    probe = title + "\n" + text[:1500]
    for name, rx in DEPT_RULES:
        m = rx.search(probe)
        if m and m.start() < best:
            dept, best = name, m.start()
    if not dept and FOREX_TOPICS & set(topics):
        dept = "FED - Foreign Exchange"
    if not dept and category == "Press_Releases":
        dept = "Press Release"
    ref = REF_PATTERN.search(text)
    return topics, dept or "Not identified", ref.group(0).strip()[:120] if ref else ""


def classify_files(title, category, files):
    """Classify using cached results where the PDF has not changed."""
    try:
        cache = json.loads(CLASS_CACHE.read_text(encoding="utf-8")) if CLASS_CACHE.exists() else {}
    except Exception:
        cache = {}
    main = files[0]
    key = str(main.relative_to(BASE_DIR))
    stamp = f"{CLASS_VERSION}:{main.stat().st_mtime_ns}"
    hit = cache.get(key)
    if hit and hit.get("v") == stamp:
        return hit["g"], hit["dp"], hit["r"]
    text = pdf_text(main, max_pages=2)
    g, dp, r = classify(title, text, category)
    cache[key] = {"v": stamp, "g": g, "dp": dp, "r": r}
    CLASS_CACHE.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    return g, dp, r


# ---------------------------------------------------------------- dashboard
def build_dashboard():
    from urllib.parse import quote
    try:
        idx = json.loads(INDEX_FILE.read_text(encoding="utf-8")) if INDEX_FILE.exists() else {}
    except Exception:
        idx = {}
    items = {}
    for p in BASE_DIR.rglob("*"):
        if p.suffix.lower() != ".pdf" or not p.is_file():
            continue
        parts = p.relative_to(BASE_DIR).parts
        if len(parts) < 4 or not re.match(r"FY \d{4}-\d{2}$", parts[0]) \
                or not re.match(r"\d{4}-\d{2}-\d{2}$", parts[1]):
            continue
        fy, d, cat = parts[0], parts[1], parts[2]
        title, sep, fname = p.name.rpartition("__")
        if not sep:
            title, fname = p.stem, p.name
        key = f"{d}|{cat}|{title}"
        meta = idx.get(key, {})
        it = items.setdefault(key, {"t": fix_text(meta.get("t") or title), "d": d, "fy": fy[3:],
                                    "c": cat.replace("_", " "), "u": meta.get("u", ""),
                                    "r": meta.get("r", ""), "f": []})
        it["f"].append({"n": fname, "h": "/".join(quote(x) for x in parts),
                        "s": p.stat().st_size, "p": p})
    data = sorted(items.values(), key=lambda i: (i["d"], i["t"]), reverse=True)
    total = len(data)
    for n, it in enumerate(data, 1):
        it["f"].sort(key=lambda f: (len(f["n"]), f["n"]))
        cat = it["c"].replace(" ", "_")
        try:
            g, dp, r = classify_files(it["t"], cat, [f["p"] for f in it["f"]])
        except Exception as e:
            logging.warning("Could not classify %s: %s", it["t"], e)
            g, dp, r = classify(it["t"], "", cat)
        it["g"], it["dp"] = g, dp
        it["r"] = it["r"] or r
        for f in it["f"]:
            del f["p"]
        if n % 50 == 0:
            print(f"  classified {n} of {total}...")
    meta = {"built": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"), "base": str(BASE_DIR)}
    page = (DASHBOARD_TEMPLATE
            .replace("__DATA__", json.dumps(data, ensure_ascii=False).replace("</", "<\\/"))
            .replace("__META__", json.dumps(meta, ensure_ascii=False).replace("</", "<\\/")))
    DASHBOARD_FILE.write_text(page, encoding="utf-8")
    logging.info("Dashboard updated: %s (%d publications)", DASHBOARD_FILE, len(data))
    return len(data)


# ---------------------------------------------------------------- main
def check_connection():
    try:
        r = session.get("https://www.rbi.org.in/", timeout=60)
        r.raise_for_status()
        home = BeautifulSoup(r.content, "html.parser")
        GLOBAL_PDFS.update(urljoin(r.url, a["href"]) for a in home.find_all("a", href=True)
                           if a["href"].lower().split("?")[0].endswith(".pdf"))
        return True
    except Exception as e:
        logging.error("Cannot connect to rbi.org.in: %s", e)
        return False


def main():
    import sys
    if "--dashboard" in sys.argv:
        n = build_dashboard()
        print(f"Dashboard created with {n} publications:\n  {DASHBOARD_FILE}")
        if os.name == "nt":
            os.startfile(DASHBOARD_FILE)
        return

    if "--fix-names" in sys.argv:            # rename files with garbled characters
        count = 0
        for p in sorted(BASE_DIR.rglob("*"), key=lambda x: -len(x.parts)):
            new = fix_text(p.name)
            if new != p.name and not (p.parent / new).exists():
                p.rename(p.parent / new)
                count += 1
        print(f"Renamed {count} file(s).")
        build_dashboard()
        return

    if "--test-email" in sys.argv:            # send the newest saved PDF as a test
        pdfs = sorted((p for p in BASE_DIR.rglob("*") if p.suffix.lower() == ".pdf" and is_real_pdf(p)),
                      key=lambda p: p.stat().st_mtime)
        if not pdfs:
            print("No PDFs found to test with.")
            return
        f = pdfs[-1]
        d = date.fromisoformat(f.parent.parent.name) if re.match(r"\d{4}-\d{2}-\d{2}$", f.parent.parent.name) else date.today()
        item = {"category": f.parent.name, "title": f.name.split("__")[0], "url": "https://www.rbi.org.in/",
                "date": d, "files": [f], "text": "", "ref": ""}
        send_emails([item])
        print("Test email sent (check the log if it did not arrive).")
        return

    if not check_connection():
        print("Could not connect to the RBI website (proxy/firewall?).")
        raise SystemExit(1)

    if "--repair" in sys.argv:               # delete broken files and re-download
        bad = [p for p in BASE_DIR.rglob("*") if p.suffix.lower() == ".pdf" and not is_real_pdf(p)]
        for p in bad:
            p.unlink()
        done = load_state()                  # items already verified are skipped (resumable)
        print(f"Removed {len(bad)} broken file(s). Already verified: {len(done)}. "
              "Checking the rest - no emails will be sent. Safe to stop and restart.")
        logging.info("Repair: removed %d broken files", len(bad))
        try:
            from_rss(done)
            from_listing_pages(done)
        except RBIBlocking:
            save_state(done)
            print("\nRBI's server is refusing connections right now. Progress is saved.\n"
                  "Wait 30-60 minutes and run  python rbi_downloader.py --repair  again.")
            return
        save_state(done)
        BASELINE_FILE.write_text(f"Baseline completed {datetime.now():%d-%b-%Y %H:%M}\n")
        still_bad = sum(1 for p in BASE_DIR.rglob("*") if p.suffix.lower() == ".pdf" and not is_real_pdf(p))
        print(f"Repair finished. Items recorded: {len(done)}. Broken files left: {still_bad}")
        build_dashboard()
        return

    baseline_done = BASELINE_FILE.exists()
    done = load_state()
    blocked = False
    try:
        from_rss(done)
        from_listing_pages(done)
    except RBIBlocking:
        blocked = True
        logging.warning("Stopped early: RBI refusing connections. Will resume next run.")
    save_state(done)
    logging.info("Run complete. New items: %d", len(NEW_ITEMS))
    try:
        build_dashboard()
    except Exception as e:
        logging.error("Dashboard update failed: %s", e)
    if not baseline_done:
        if not blocked:
            BASELINE_FILE.write_text(f"Baseline completed {datetime.now():%d-%b-%Y %H:%M}\n")
            logging.info("Baseline complete - email alerts start from the next run.")
        return                               # never email the historical backlog
    if NEW_ITEMS and EMAIL_ENABLED:
        send_emails(NEW_ITEMS)


DASHBOARD_TEMPLATE = r'''<!DOCTYPE html>
<html lang="en-IN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>RBI Regulatory Register</title>
<style>
:root{
  --bg:#F3F5F7; --surface:#FFFFFF; --sunk:#E9EDF1; --ink:#17212B; --ink-2:#4A5866; --ink-3:#7C8894;
  --line:#D9DFE5; --brand:#0F3D56; --brand-ink:#FFFFFF; --gold:#B8862B; --focus:#2F7CC0;
  --c-notifications:#0F7A82; --c-master-directions:#4152A3; --c-master-circulars:#B0761A; --c-press-releases:#6B7785; --c-other:#8A6FB0;
  --serif:"Iowan Old Style","Charter","Cambria",Georgia,serif;
  --sans:"Segoe UI Variable Text","Segoe UI",system-ui,-apple-system,Roboto,Arial,sans-serif;
  color-scheme:light;
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#10171E; --surface:#17212B; --sunk:#1E2A36; --ink:#E6ECF1; --ink-2:#AAB7C3; --ink-3:#7D8B98;
    --line:#2A3846; --brand:#16506F; --brand-ink:#FFFFFF; --gold:#D8A64A; --focus:#6AB0EE;
    --c-notifications:#2FB3BA; --c-master-directions:#8592E0; --c-master-circulars:#E0A746; --c-press-releases:#97A3B0; --c-other:#B79BDB;
    color-scheme:dark;
  }
}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--bg);color:var(--ink);font:15px/1.5 var(--sans);-webkit-font-smoothing:antialiased}
a{color:inherit}
:focus-visible{outline:2px solid var(--focus);outline-offset:2px;border-radius:4px}
.num{font-variant-numeric:tabular-nums}

/* masthead */
.mast{background:var(--brand);color:var(--brand-ink)}
.mast-in{max-width:1320px;margin:0 auto;padding:28px 28px 0}
.mast-top{display:flex;justify-content:space-between;align-items:flex-end;gap:24px;flex-wrap:wrap}
.mast h1{font:600 30px/1.15 var(--serif);margin:0;letter-spacing:-.01em}
.mast p{margin:6px 0 0;color:rgba(255,255,255,.78);font-size:14px;max-width:72ch}
.stats{display:flex;gap:28px;flex-wrap:wrap}
.stat b{display:block;font:600 24px/1.1 var(--serif)}
.stat span{font-size:12.5px;color:rgba(255,255,255,.72)}

/* timeline: the signature element */
.tl{margin-top:22px;padding:14px 0 0;border-top:1px solid rgba(255,255,255,.14)}
.tl-head{display:flex;justify-content:space-between;align-items:baseline;font-size:13px;color:rgba(255,255,255,.8);margin-bottom:8px}
.tl-head em{font-style:normal;color:#fff}
.tl-bars{display:grid;grid-template-columns:repeat(24,1fr);gap:4px;align-items:end;height:96px}
.tl-col{position:relative;height:100%;display:flex;flex-direction:column;justify-content:flex-end;border:0;background:none;padding:0;cursor:pointer;color:inherit;font:inherit}
.tl-col .stack{display:flex;flex-direction:column-reverse;border-radius:3px 3px 0 0;overflow:hidden;min-height:2px;background:rgba(255,255,255,.12)}
.tl-col .stack i{display:block}
.tl-col:hover .stack{filter:brightness(1.15)}
.tl-col.sel .stack{outline:2px solid var(--gold);outline-offset:2px}
.tl-col.dim{opacity:.35}
.tl-col.fy::before{content:"";position:absolute;left:-3px;top:0;bottom:0;border-left:1px dashed rgba(255,255,255,.35)}
.tl-col .fyl{position:absolute;top:-2px;left:3px;font-size:10.5px;color:rgba(255,255,255,.6);white-space:nowrap}
.tl-labels{display:grid;grid-template-columns:repeat(24,1fr);gap:4px;font-size:10.5px;color:rgba(255,255,255,.6);padding:6px 0 14px;text-align:center}
.legend{display:flex;gap:16px;flex-wrap:wrap;font-size:12.5px;color:rgba(255,255,255,.82);padding-bottom:16px}
.legend span::before{content:"";display:inline-block;width:9px;height:9px;border-radius:2px;margin-right:6px;background:var(--dot)}

/* layout */
.wrap{max-width:1320px;margin:0 auto;padding:24px 28px 60px;display:grid;grid-template-columns:250px 1fr;gap:28px}
aside{position:sticky;top:16px;align-self:start}
.panel{margin-bottom:24px}
.panel h2{font:600 13px/1.3 var(--sans);color:var(--ink-2);margin:0 0 8px}
.opt{display:flex;justify-content:space-between;align-items:center;width:100%;padding:7px 10px;margin:1px 0;border:0;border-radius:6px;background:none;color:var(--ink);font:inherit;font-size:14px;cursor:pointer;text-align:left}
.opt:hover{background:var(--sunk)}
.opt.on{background:var(--surface);box-shadow:inset 0 0 0 1px var(--line);font-weight:600}
.opt .cnt{color:var(--ink-3);font-size:12.5px}
.opt .dot{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:9px;background:var(--dot);vertical-align:-1px}
.opt.off{color:var(--ink-3)}
.opt.off .dot{background:transparent;box-shadow:inset 0 0 0 1.5px var(--dot)}
.seg{display:grid;grid-template-columns:repeat(3,1fr);background:var(--sunk);border-radius:8px;padding:3px}
.seg button{border:0;background:none;padding:6px 4px;border-radius:6px;font:inherit;font-size:13px;color:var(--ink-2);cursor:pointer}
.seg button.on{background:var(--surface);color:var(--ink);font-weight:600;box-shadow:0 1px 2px rgba(0,0,0,.08)}
.opt.sub{position:relative;overflow:hidden}
.opt.sub .lbl{position:relative;z-index:1}
.opt.sub .cnt{position:relative;z-index:1}
.opt.sub .meter{position:absolute;left:0;top:3px;bottom:3px;border-radius:5px;background:var(--sunk);z-index:0}
.opt.sub:hover .meter{background:var(--line)}
.opt.sub.on .meter{background:transparent}
.scroll{max-height:340px;overflow:auto;margin-right:-6px;padding-right:6px}
.tags{display:flex;gap:6px;flex-wrap:wrap;margin-top:8px}
.tag{border:0;padding:2px 9px;border-radius:999px;background:var(--sunk);color:var(--ink-2);font:500 12px/1.6 var(--sans);cursor:pointer}
.tag:hover{background:var(--brand);color:#fff}
.tag.dept{background:transparent;box-shadow:inset 0 0 0 1px var(--line)}
.foot{font-size:12px;color:var(--ink-3);line-height:1.5;word-break:break-word}

/* toolbar */
.bar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-bottom:14px}
.search{flex:1;min-width:240px;position:relative}
.search input{width:100%;padding:11px 14px 11px 38px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);font:inherit}
.search svg{position:absolute;left:12px;top:50%;transform:translateY(-50%);color:var(--ink-3)}
.count{color:var(--ink-2);font-size:14px}
.chips{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:6px}
.chip{display:inline-flex;align-items:center;gap:6px;padding:4px 6px 4px 10px;border-radius:999px;background:var(--surface);box-shadow:inset 0 0 0 1px var(--line);font-size:13px}
.chip button{border:0;background:var(--sunk);color:var(--ink-2);width:20px;height:20px;border-radius:50%;cursor:pointer;font-size:13px;line-height:1}
.linkbtn{border:0;background:none;color:var(--focus);font:inherit;font-size:13px;cursor:pointer;padding:4px}

/* list */
.day{margin-top:22px}
.day h3{position:sticky;top:0;z-index:1;background:var(--bg);margin:0;padding:8px 0;font:600 15px/1.3 var(--serif);display:flex;justify-content:space-between;border-bottom:1px solid var(--line)}
.day h3 small{font:400 12.5px var(--sans);color:var(--ink-3)}
.row{display:grid;grid-template-columns:4px 1fr auto;gap:0 16px;padding:14px 14px 14px 0;border-bottom:1px solid var(--line);background:var(--surface)}
.row:first-of-type{border-top:0}
.rail{background:var(--cat);border-radius:0 2px 2px 0}
.row h4{margin:0;font:600 15px/1.4 var(--sans);max-width:80ch}
.row h4 a{text-decoration:none}
.row h4 a:hover{text-decoration:underline}
.meta{margin-top:4px;display:flex;gap:14px;flex-wrap:wrap;font-size:12.5px;color:var(--ink-3)}
.meta .cat{color:var(--cat);font-weight:600}
.new{display:inline-block;margin-left:8px;padding:1px 7px;border-radius:4px;background:var(--gold);color:#fff;font:600 11px/1.6 var(--sans);vertical-align:2px}
.files{display:flex;gap:6px;flex-wrap:wrap;justify-content:flex-end;align-content:flex-start;max-width:330px}
.pdf{display:inline-flex;align-items:center;gap:6px;padding:6px 10px;border-radius:6px;background:var(--sunk);text-decoration:none;font-size:12.5px;white-space:nowrap}
.pdf:hover{background:var(--brand);color:#fff}
.pdf svg{flex:none}
.pdf small{color:inherit;opacity:.7}
.more{display:block;margin:24px auto 0;padding:10px 22px;border:1px solid var(--line);border-radius:8px;background:var(--surface);color:var(--ink);font:inherit;cursor:pointer}
.empty{padding:60px 20px;text-align:center;color:var(--ink-2)}
.empty b{display:block;font:600 18px var(--serif);color:var(--ink);margin-bottom:6px}

@media (max-width:900px){
  .wrap{grid-template-columns:1fr;padding:18px 16px 48px}
  aside{position:static;display:grid;grid-template-columns:1fr 1fr;gap:0 18px}
  .mast-in{padding:20px 16px 0}
  .tl-labels span:nth-child(odd){visibility:hidden}
  .row{grid-template-columns:4px 1fr}
  .files{grid-column:2;justify-content:flex-start;margin-top:10px;max-width:none}
}
@media (max-width:560px){aside{grid-template-columns:1fr}.mast h1{font-size:24px}}
@media (prefers-reduced-motion:no-preference){.tl-col .stack{transition:filter .15s}}
</style>
</head>
<body>
<header class="mast">
  <div class="mast-in">
    <div class="mast-top">
      <div>
        <h1>RBI Regulatory Register</h1>
        <p id="sub"></p>
      </div>
      <div class="stats" id="stats"></div>
    </div>
    <div class="tl">
      <div class="tl-head"><span>Publications by month, last 24 months. Select a month to filter.</span><em id="tlsel"></em></div>
      <div class="tl-bars" id="bars"></div>
      <div class="tl-labels" id="labels"></div>
      <div class="legend" id="legend"></div>
    </div>
  </div>
</header>

<div class="wrap">
  <aside>
    <div class="panel">
      <h2>Period</h2>
      <div class="seg" id="period">
        <button data-p="all" class="on">All</button><button data-p="30">30 days</button><button data-p="7">7 days</button>
      </div>
    </div>
    <div class="panel"><h2>Type</h2><div id="cats"></div></div>
    <div class="panel"><h2>Subject</h2><div class="scroll" id="topics"></div></div>
    <div class="panel"><h2>Issuing department</h2><div class="scroll" id="depts"></div></div>
    <div class="panel"><h2>Financial year</h2><div id="fys"></div></div>
    <div class="panel foot" id="foot"></div>
  </aside>

  <main>
    <div class="bar">
      <label class="search">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
        <input id="q" type="search" placeholder="Search titles, subjects, reference numbers or file names" aria-label="Search publications">
      </label>
      <span class="count num" id="count"></span>
    </div>
    <div class="chips" id="chips"></div>
    <div id="list"></div>
  </main>
</div>

<script id="data" type="application/json">__DATA__</script>
<script id="meta" type="application/json">__META__</script>
<script>
(function(){
  const DATA = JSON.parse(document.getElementById('data').textContent);
  const META = JSON.parse(document.getElementById('meta').textContent);
  const ORDER = ['Notifications','Master Directions','Master Circulars','Press Releases'];
  const slug = c => ORDER.includes(c) ? c.toLowerCase().replace(/ /g,'-') : 'other';
  const cvar = c => 'var(--c-' + slug(c) + ')';
  const esc = s => String(s).replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
  const pd = s => { const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); };
  const built = new Date(META.built);
  const DAY = 864e5;
  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const MONL = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const fmtSize = b => b > 1048576 ? (b/1048576).toFixed(1)+' MB' : Math.max(1, Math.round(b/1024))+' KB';
  const fmtDay = d => d.toLocaleDateString('en-IN',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

  DATA.forEach(it => { it.dt = pd(it.d); it.m = it.d.slice(0,7);
    it.g = it.g || ['Others']; it.dp = it.dp || 'Not identified';
    it.s = (it.t + ' ' + it.r + ' ' + it.g.join(' ') + ' ' + it.dp + ' ' + it.f.map(f => f.n).join(' ')).toLowerCase(); });

  const cats = [...new Set(DATA.map(i => i.c))].sort((a,b) => {
    const ia = ORDER.indexOf(a), ib = ORDER.indexOf(b);
    return (ia<0?99:ia) - (ib<0?99:ib) || a.localeCompare(b); });
  const fys = [...new Set(DATA.map(i => i.fy))].sort().reverse();
  const curFY = (() => { const y = built.getMonth() >= 3 ? built.getFullYear() : built.getFullYear()-1;
    return y + '-' + String(y+1).slice(2); })();

  const st = { q:'', fy:'all', off:new Set(), period:'all', month:null, topic:'all', dept:'all', limit:120 };
  const countBy = f => { const m = {}; DATA.forEach(i => [].concat(f(i)).forEach(k => m[k] = (m[k]||0) + 1));
    return Object.entries(m).sort((a,b) => (a[0]==='Others'||a[0]==='Not identified') - (b[0]==='Others'||b[0]==='Not identified') || b[1]-a[1] || a[0].localeCompare(b[0])); };

  /* masthead */
  const files = DATA.reduce((n,i) => n + i.f.length, 0);
  const bytes = DATA.reduce((n,i) => n + i.f.reduce((a,f) => a + f.s, 0), 0);
  const last30 = DATA.filter(i => built - i.dt <= 30*DAY).length;
  document.getElementById('sub').textContent =
    'Circulars, notifications, Master Directions and press releases from the Reserve Bank of India, filed by financial year and date. Updated ' +
    built.toLocaleString('en-IN',{day:'numeric',month:'short',year:'numeric',hour:'numeric',minute:'2-digit'}) + '.';
  document.getElementById('stats').innerHTML = [
    [DATA.length.toLocaleString('en-IN'), 'publications'],
    [DATA.filter(i => i.fy === curFY).length, 'in FY ' + curFY],
    [last30, 'in the last 30 days'],
    [files.toLocaleString('en-IN'), 'PDFs, ' + fmtSize(bytes)]
  ].map(([b,s]) => '<div class="stat"><b class="num">'+b+'</b><span>'+esc(s)+'</span></div>').join('');

  /* timeline */
  const months = [];
  for (let k = 23; k >= 0; k--) { const d = new Date(built.getFullYear(), built.getMonth()-k, 1);
    months.push({ key: d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0'), d }); }
  const byMonth = {}; DATA.forEach(i => { (byMonth[i.m] = byMonth[i.m] || []).push(i); });
  const maxM = Math.max(1, ...months.map(m => (byMonth[m.key]||[]).length));
  const bars = document.getElementById('bars');
  bars.innerHTML = months.map(m => {
    const list = byMonth[m.key] || [];
    const segs = cats.map(c => { const n = list.filter(i => i.c === c).length;
      return n ? '<i style="height:'+(n/list.length*100)+'%;background:'+cvar(c)+'"></i>' : ''; }).join('');
    const h = list.length ? Math.max(4, list.length/maxM*100) : 2;
    const fy = m.d.getMonth() === 3;
    const label = MONL[m.d.getMonth()] + ' ' + m.d.getFullYear() + ': ' + list.length + ' publication' + (list.length===1?'':'s');
    return '<button class="tl-col'+(fy?' fy':'')+'" data-m="'+m.key+'" title="'+label+'" aria-label="'+label+'">' +
      (fy ? '<span class="fyl">FY '+m.d.getFullYear()+'-'+String(m.d.getFullYear()+1).slice(2)+'</span>' : '') +
      '<span class="stack" style="height:'+h+'%">'+segs+'</span></button>';
  }).join('');
  document.getElementById('labels').innerHTML = months.map(m =>
    '<span>'+MON[m.d.getMonth()]+(m.d.getMonth()===0?' ’'+String(m.d.getFullYear()).slice(2):'')+'</span>').join('');
  document.getElementById('legend').innerHTML = cats.map(c => '<span style="--dot:'+cvar(c)+'">'+esc(c)+'</span>').join('');
  bars.addEventListener('click', e => { const b = e.target.closest('.tl-col'); if (!b) return;
    st.month = st.month === b.dataset.m ? null : b.dataset.m; st.limit = 120; render(); });

  /* sidebar */
  const catBox = document.getElementById('cats'), fyBox = document.getElementById('fys');
  catBox.innerHTML = cats.map(c => '<button class="opt" data-c="'+esc(c)+'" style="--dot:'+cvar(c)+'" aria-pressed="true"><span><span class="dot"></span>'+esc(c)+'</span><span class="cnt num">'+DATA.filter(i=>i.c===c).length+'</span></button>').join('');
  fyBox.innerHTML = '<button class="opt on" data-fy="all"><span>All years</span><span class="cnt num">'+DATA.length+'</span></button>' +
    fys.map(f => '<button class="opt" data-fy="'+f+'"><span>FY '+f+'</span><span class="cnt num">'+DATA.filter(i=>i.fy===f).length+'</span></button>').join('');
  const topBox = document.getElementById('topics'), depBox = document.getElementById('depts');
  const listOpts = (box, entries, key, allLabel) => {
    const max = Math.max(1, ...entries.map(e => e[1]));
    box.innerHTML = '<button class="opt sub on" data-'+key+'="all"><span class="lbl">'+allLabel+'</span><span class="cnt num">'+DATA.length+'</span></button>' +
      entries.map(([k,n]) => '<button class="opt sub" data-'+key+'="'+esc(k)+'"><span class="meter" style="width:'+(n/max*100)+'%"></span><span class="lbl">'+esc(k)+'</span><span class="cnt num">'+n+'</span></button>').join('');
  };
  listOpts(topBox, countBy(i => i.g), 'topic', 'All subjects');
  listOpts(depBox, countBy(i => i.dp), 'dept', 'All departments');
  topBox.addEventListener('click', e => { const b = e.target.closest('.opt'); if (!b) return;
    st.topic = b.dataset.topic; st.limit = 120; render(); });
  depBox.addEventListener('click', e => { const b = e.target.closest('.opt'); if (!b) return;
    st.dept = b.dataset.dept; st.limit = 120; render(); });
  catBox.addEventListener('click', e => { const b = e.target.closest('.opt'); if (!b) return;
    const c = b.dataset.c; st.off.has(c) ? st.off.delete(c) : st.off.add(c);
    if (st.off.size === cats.length) st.off.clear(); st.limit = 120; render(); });
  fyBox.addEventListener('click', e => { const b = e.target.closest('.opt'); if (!b) return;
    st.fy = b.dataset.fy; st.limit = 120; render(); });
  document.getElementById('period').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return;
    st.period = b.dataset.p; st.limit = 120; render(); });
  let tmr; document.getElementById('q').addEventListener('input', e => { clearTimeout(tmr);
    tmr = setTimeout(() => { st.q = e.target.value.trim().toLowerCase(); st.limit = 120; render(); }, 120); });
  document.getElementById('foot').innerHTML = 'Files are stored in<br>'+esc(META.base)+'<br><br>Run the downloader to refresh this page with new publications.';

  function filtered(){
    const words = st.q.split(/\s+/).filter(Boolean);
    return DATA.filter(i =>
      !st.off.has(i.c) && (st.fy === 'all' || i.fy === st.fy) && (!st.month || i.m === st.month) &&
      (st.topic === 'all' || i.g.includes(st.topic)) && (st.dept === 'all' || i.dp === st.dept) &&
      (st.period === 'all' || built - i.dt <= (+st.period) * DAY) && words.every(w => i.s.includes(w)));
  }

  const pdfIcon = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9z"/><path d="M14 3v6h6"/></svg>';

  function render(){
    const res = filtered();
    document.getElementById('count').textContent = res.length.toLocaleString('en-IN') + ' of ' + DATA.length.toLocaleString('en-IN') + ' shown';
    catBox.querySelectorAll('.opt').forEach(b => { const off = st.off.has(b.dataset.c);
      b.classList.toggle('off', off); b.setAttribute('aria-pressed', String(!off)); });
    fyBox.querySelectorAll('.opt').forEach(b => b.classList.toggle('on', b.dataset.fy === st.fy));
    topBox.querySelectorAll('.opt').forEach(b => b.classList.toggle('on', b.dataset.topic === st.topic));
    depBox.querySelectorAll('.opt').forEach(b => b.classList.toggle('on', b.dataset.dept === st.dept));
    document.querySelectorAll('#period button').forEach(b => b.classList.toggle('on', b.dataset.p === st.period));
    bars.querySelectorAll('.tl-col').forEach(b => { b.classList.toggle('sel', b.dataset.m === st.month);
      b.classList.toggle('dim', !!st.month && b.dataset.m !== st.month); });
    const ml = st.month ? MONL[+st.month.slice(5)-1] + ' ' + st.month.slice(0,4) : '';
    document.getElementById('tlsel').textContent = ml ? 'Showing ' + ml : '';

    const chips = [];
    if (st.month) chips.push(['month', ml]);
    if (st.fy !== 'all') chips.push(['fy', 'FY ' + st.fy]);
    if (st.topic !== 'all') chips.push(['topic', st.topic]);
    if (st.dept !== 'all') chips.push(['dept', st.dept]);
    if (st.period !== 'all') chips.push(['period', 'Last ' + st.period + ' days']);
    if (st.off.size) chips.push(['cats', 'Hiding ' + [...st.off].join(', ')]);
    if (st.q) chips.push(['q', '“' + st.q + '”']);
    document.getElementById('chips').innerHTML = chips.map(([k,l]) =>
      '<span class="chip">'+esc(l)+'<button data-k="'+k+'" aria-label="Remove filter '+esc(l)+'">×</button></span>').join('') +
      (chips.length > 1 ? '<button class="linkbtn" data-k="all">Clear all filters</button>' : '');

    const list = document.getElementById('list');
    if (!res.length) { list.innerHTML = '<div class="empty"><b>No publications match these filters</b>Try a different search term, or <button class="linkbtn" data-k="all">clear all filters</button>.</div>'; return; }
    const shown = res.slice(0, st.limit);
    let html = '', cur = '';
    shown.forEach(i => {
      if (i.d !== cur) {
        if (cur) html += '</section>';
        cur = i.d;
        const n = res.filter(x => x.d === cur).length;
        html += '<section class="day"><h3><span>'+fmtDay(i.dt)+'</span><small class="num">'+n+' publication'+(n===1?'':'s')+'</small></h3>';
      }
      const isNew = built - i.dt <= 7*DAY;
      const title = i.u ? '<a href="'+esc(i.u)+'" target="_blank" rel="noopener">'+esc(i.t)+'</a>' : esc(i.t);
      const files = i.f.map((f,k) => '<a class="pdf" href="'+f.h+'" target="_blank" title="'+esc(f.n)+'">'+pdfIcon+
        (i.f.length === 1 ? 'Open PDF' : (k === 0 ? 'Main PDF' : 'Part ' + (k+1)))+' <small class="num">'+fmtSize(f.s)+'</small></a>').join('');
      html += '<article class="row" style="--cat:'+cvar(i.c)+'"><span class="rail"></span><div><h4>'+title+(isNew?'<span class="new">New</span>':'')+'</h4>' +
        '<div class="meta"><span class="cat">'+esc(i.c)+'</span><span>FY '+esc(i.fy)+'</span>'+(i.r?'<span>'+esc(i.r)+'</span>':'')+'</div>' +
        '<div class="tags">'+i.g.map(g => '<button class="tag" data-settopic="'+esc(g)+'">'+esc(g)+'</button>').join('') +
        (i.dp !== 'Not identified' ? '<button class="tag dept" data-setdept="'+esc(i.dp)+'">'+esc(i.dp)+'</button>' : '') + '</div></div>' +
        '<div class="files">'+files+'</div></article>';
    });
    html += '</section>';
    if (res.length > shown.length) html += '<button class="more" id="more">Show '+Math.min(120, res.length-shown.length)+' more of '+(res.length-shown.length).toLocaleString('en-IN')+' remaining</button>';
    list.innerHTML = html;
    const more = document.getElementById('more');
    if (more) more.onclick = () => { st.limit += 120; render(); };
  }

  document.addEventListener('click', e => { const t = e.target.closest('[data-settopic],[data-setdept]'); if (!t) return;
    if (t.dataset.settopic) st.topic = t.dataset.settopic; else st.dept = t.dataset.setdept;
    st.limit = 120; render(); window.scrollTo({top: document.querySelector('.wrap').offsetTop - 8}); });
  document.addEventListener('click', e => { const b = e.target.closest('[data-k]'); if (!b) return;
    const k = b.dataset.k;
    if (k === 'month' || k === 'all') st.month = null;
    if (k === 'fy' || k === 'all') st.fy = 'all';
    if (k === 'period' || k === 'all') st.period = 'all';
    if (k === 'cats' || k === 'all') st.off.clear();
    if (k === 'topic' || k === 'all') st.topic = 'all';
    if (k === 'dept' || k === 'all') st.dept = 'all';
    if (k === 'q' || k === 'all') { st.q = ''; document.getElementById('q').value = ''; }
    st.limit = 120; render(); });

  render();
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
