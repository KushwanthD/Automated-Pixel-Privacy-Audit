"""
Healthcare Pixel Privacy Auditor
══════════════════════════════════════════════════════════════════════════════
Workflow per domain
  1. Discover pages  — crawl the homepage for internal links; classify by type
                       (portal, scheduling, telehealth, login, intake, etc.).
                       Probe common paths for any type not found via crawl.
  2. Clean visit     — open every page as a first-time visitor in a fresh
                       browser session (no prior cookies). Do NOT interact with
                       consent banners.  Record every outbound network request.
  3. Classify timing — PRE-BANNER   : tracker fires before consent UI exists
                       BANNER-PRESENT: tracker fires after banner text appears
                                       but before any click
                       NO TRACKER   : nothing fires during the untouched session
  4. Payload check   — inspect each tracker request URL for page-path, event
                       names, PII parameters, fingerprint cookies, and health-
                       context keywords that indicate sensitive data exposure.
  5. Risk score      — combine page-type base risk (portal=10, scheduling=9 …)
                       with timing penalty and payload flag count → 1-10 score.
  6. Excel report    — three sheets: Executive Summary / Page Findings /
                       Tracker Detail, all colour-coded by risk level.
══════════════════════════════════════════════════════════════════════════════
"""

import re
import subprocess
import sys
import os
from datetime import datetime
from urllib.parse import urlparse, unquote

# ── Auto-install dependencies ─────────────────────────────────────────────────
for _pkg in ["playwright", "openpyxl"]:
    try:
        __import__(_pkg)
    except ImportError:
        subprocess.run([sys.executable, "-m", "pip", "install", "-q", _pkg])

subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ═══════════════════════════════════════════════════════════════════════════════
# USER CONFIGURATION
# Paste domains / URLs below — one per line.  Mix bare domains, full URLs,
# Google Maps links, Google Ads /aclk links, phone numbers, and general noise;
# the parser strips everything that is not a real business domain.
# ═══════════════════════════════════════════════════════════════════════════════
DOMAINS = """


"""

CITY          = "  "
BUSINESS_TYPE = "  "
# ═══════════════════════════════════════════════════════════════════════════════


# ── Ad-tech tracker vendors ───────────────────────────────────────────────────
# FIX: "Google Ads/Analytics" used to lump together four distinct Google
# products (Tag Manager container, GA4 analytics, Ads conversion tracking,
# and DoubleClick ad-serving) under one vague label. Manual verification of
# real network captures showed these need separate identification — a GTM
# container load is not the same finding as a DoubleClick ad-serving hit.
# Order matters: dicts preserve insertion order and classify_tracker() now
# returns the FIRST matching vendor, so list the more specific GA4/Ads
# patterns before the generic "doubleclick.net" catch-all.
TRACKERS = {
    "Meta/Facebook Pixel":      ("facebook.com/tr", "connect.facebook.net"),
    "Google Tag Manager":       ("googletagmanager.com/gtm.js",
                                  "googletagmanager.com/gtag/js"),
    "Google Analytics (GA4)":   ("google-analytics.com", "/g/collect",
                                  "analytics.google.com"),
    "Google Ads Conversion":    ("googleadservices.com",
                                  "googleads.g.doubleclick.net",
                                  "pagead/viewthroughconversion",
                                  "pagead/1p-user-list"),
    "Google/DoubleClick Ad-Serving": ("doubleclick.net", "googlesyndication.com",
                                  "google.com/ccm", "google.co.in/ccm",
                                  "google.com/pagead", "google.com/rmkt"),
    "TikTok Pixel":         ("analytics.tiktok.com", "ads.tiktok.com"),
    "LinkedIn Insight":     ("snap.licdn.com", "px.ads.linkedin.com"),
    "Pinterest Tag":        ("ct.pinterest.com",),
    "Snapchat Pixel":       ("tr.snapchat.com",),
    "Microsoft/Bing Ads":   ("bat.bing.com",),
    "Hotjar":               ("hotjar.com",),
    "Twitter/X Pixel":      ("analytics.twitter.com", "ads-twitter.com"),
    "Heap Analytics":       ("heapanalytics.com",),
    "Mixpanel":             ("api.mixpanel.com",),
    "Segment":              ("cdn.segment.com", "api.segment.io"),
    "Amplitude":            ("api.amplitude.com", "cdn.amplitude.com"),
    "FullStory":            ("fullstory.com", "rs.fullstory.com"),
}

# Text that signals a cookie/consent banner is present on the page
CONSENT_KEYWORDS = (
    "cookie", "consent", "accept all", "i agree", "gdpr",
    "privacy preference", "we use cookies", "cookie policy",
)

# Known consent-management-platform (CMP) script/element signatures.
# Used as a SECOND, independent signal alongside CONSENT_KEYWORDS so that a
# page is only ever classified PRE-BANNER (banner exists, tracker just beat
# it) vs NO-CONSENT-MECHANISM (no banner/CMP present at all, ever) based on
# more than text-matching alone — catching banners whose visible text might
# not contain an exact CONSENT_KEYWORDS phrase but whose CMP script is
# still present and will render a banner shortly (e.g. on slow connections).
CMP_SCRIPT_SIGNATURES = (
    "onetrust", "cookiebot", "cookieyes", "cookielaw.org",
    "trustarc", "quantcast", "didomi", "usercentrics",
    "iubenda", "termly", "complianz", "borlabs-cookie",
    "cookieconsent", "cookie-consent", "klaro",
)


# ── Healthcare page types — ordered highest → lowest risk ────────────────────
# risk: base score 1-10 used before timing / payload adjustments
PAGE_TYPES = {
    "patient_portal": {
        "label":    "Patient Portal",
        "risk":     10,
        "patterns": [r"/portal", r"/patient[._-]portal", r"/myhealth",
                     r"/my[._-]health", r"/myhealthrecord", r"/patient[._-]account",
                     r"/my[._-]account", r"/dashboard", r"/health[._-]record",
                     r"/my[._-]record"],
    },
    "appointment": {
        "label":    "Appointment / Scheduling",
        "risk":     9,
        "patterns": [r"/appointment", r"/schedul", r"/book(?:ing)?",
                     r"/request[._-]appt", r"/request[._-]appointment",
                     r"/make[._-]appt", r"/reserve"],
    },
    "telehealth": {
        "label":    "Telehealth Entry",
        "risk":     9,
        "patterns": [r"/telehealth", r"/virtual[._-]visit", r"/video[._-]visit",
                     r"/telemedicine", r"/online[._-]visit", r"/virtual[._-]care",
                     r"/e[._-]visit"],
    },
    "login": {
        "label":    "Login / Authentication",
        "risk":     8,
        "patterns": [r"/login", r"/sign[._-]?in", r"/signin",
                     r"/auth(?!/or)", r"/secure", r"/account/login", r"/sso"],
    },
    "intake_registration": {
        "label":    "Intake / Registration",
        "risk":     8,
        "patterns": [r"/intake", r"/new[._-]patient", r"/registration",
                     r"/register(?!ed)", r"/enroll", r"/get[._-]started",
                     r"/start[._-]care", r"/onboard"],
    },
    "contact_form": {
        "label":    "Contact / Inquiry Form",
        "risk":     7,
        "patterns": [r"/contact", r"/contact[._-]us", r"/reach[._-]us",
                     r"/get[._-]in[._-]touch", r"/request[._-]info"],
    },
    "condition_treatment": {
        "label":    "Condition / Treatment",
        "risk":     6,
        "patterns": [r"/condition", r"/treatment", r"/specialty",
                     r"/disease", r"/symptom", r"/therapy",
                     r"/procedure", r"/medication", r"/program"],
    },
    "provider": {
        "label":    "Provider / Doctor Page",
        "risk":     5,
        "patterns": [r"/provider", r"/doctor", r"/physician", r"/staff",
                     r"/our[._-]team", r"/meet[._-]the", r"/find[._-]a[._-]doctor",
                     r"/find[._-]provider", r"/bio/"],
    },
    "homepage": {
        "label":    "Homepage",
        "risk":     3,
        "patterns": [],   # matched by path == "/" logic below
    },
}

# Common paths to probe when link-crawl misses a page type
PROBE_PATHS = {
    "patient_portal":       ["/portal", "/patient-portal", "/my-account",
                              "/myhealth", "/myhealthrecord"],
    "appointment":          ["/appointments", "/schedule", "/book",
                              "/request-appointment", "/booking"],
    "telehealth":           ["/telehealth", "/virtual-visit", "/telemedicine"],
    "login":                ["/login", "/sign-in"],
    "intake_registration":  ["/new-patient", "/intake", "/registration"],
    "contact_form":         ["/contact", "/contact-us"],
    "condition_treatment":  ["/services", "/conditions", "/specialties"],
    "provider":             ["/providers", "/our-team", "/doctors"],
}


# ── Payload sensitivity patterns ──────────────────────────────────────────────
# Each entry: (regex_pattern, human_readable_flag_label, severity)
#   severity "high"   → real personal identifiers (email, phone, name, address)
#   severity "medium"  → browser/device tracking identifiers (fbp, auid, ttclid)
#   severity "info"    → contextual signals (health keywords, URLs, cookies)
#
# FIX (confirmed against real captured payloads):
#   The old PII pattern `(uid|user_id|external_id|em=|ph=|fn=|ln=|db=|ge=|ct=|st=|zp=)`
#   matched as a bare substring anywhere in the URL. On real Tebra.com traffic
#   this produced false positives on EVERY page:
#     - "auid=15125467..."  (Google's anonymous ad-user-id) matched "uid"
#     - "fst=1782500019991" (Google's first-seen-timestamp) matched "st="
#   Neither is a real identifier. Fixed by anchoring every key to an actual
#   query-string position: `[?&]key=` (start of param) so "auid" can never
#   match the bare key "uid", and "fst=" can never match "st=".
#
#   The conversion-event pattern was already correctly anchored on `ev=` in
#   the prior version and is kept as-is — verified it does NOT fire on
#   PageView / ViewContent, only on genuine conversion event names.
# ─────────────────────────────────────────────────────────────────────────────
SENSITIVE_PATTERNS = [
    (r"[?&](dl|cd\[page_location\]|ep\.page_location|page_location|document_location)=([^&]{4,})",
     "Page URL/path transmitted in pixel payload", "info"),
    (r"[?&](content_name|cd\[content_name\])=([^&]{3,})",
     "Content name transmitted in pixel payload", "info"),
    (r"\b(appointment|schedule|book|portal|intake|register|login|telehealth"
     r"|symptom|diagnosis|treatment|medication|prescription|mental.health"
     r"|therapy|psychiatry|oncology|fertility)\b",
     "Health-context keyword found in tracker request URL", "info"),
    # Anchored on the ev= key — never matches path words like "schedule".
    (r"[?&]ev=(Lead|CompleteRegistration|Schedule|Contact"
     r"|InitiateCheckout|Purchase|SubmitApplication|Subscribe"
     r"|StartTrial|AddPaymentInfo)(?:&|$)",
     "Sensitive conversion event fired (Lead / Schedule / Registration)", "high"),
    # ── HIGH severity: genuine personal identifiers, anchored to param keys ──
    # em/ph/fn/ln/db/ge/ct/st/zp are Meta's Advanced Matching fields (usually
    # SHA-256 hashed, but still classified as PII parameters per FTC guidance
    # on hashed health-related identifiers). external_id / user_id are
    # site-assigned identity keys, not generic browser IDs.
    (r"[?&](em|ph|fn|ln|db|ge|ct|st|zp|external_id|user_id)=([^&]{2,})",
     "Likely PII parameter detected (name/email/phone/address/identity field)", "high"),
    # ── MEDIUM severity: browser/device tracking identifiers, not personal
    # data. Phrasing follows the precise, defensible pattern: name the exact
    # parameter and destination, and state explicitly that it is a
    # persistent browser/device identifier rather than a personal identifier
    # — e.g. "Persistent browser identifier (_fbp) transmitted to Meta. No
    # direct user identifiers (email/phone/name) observed in this request."
    (r"[?&](fbp|_fbp|fbc|_fbc)=([^&]{2,})",
     "Persistent browser identifier (fbp/fbc) transmitted to Meta. "
     "No direct user identifiers (email/phone/name) observed in this request.", "medium"),
    (r"[?&](_ga|_gid|_gat|_gcl_au|auid)=([^&]{2,})",
     "Persistent browser/ad identifier transmitted to Google. "
     "No direct user identifiers (email/phone/name) observed in this request.", "medium"),
    (r"[?&](ttclid|ttp|_ttp)=([^&]{2,})",
     "Persistent click identifier transmitted to TikTok. "
     "No direct user identifiers (email/phone/name) observed in this request.", "medium"),
    (r"[?&](uid|uuid|client_id|cid)=([^&]{2,})",
     "Persistent client/session identifier transmitted. "
     "No direct user identifiers (email/phone/name) observed in this request.", "medium"),
    (r"[?&](form|field|input|value)=([^&]{3,})",
     "Possible form-field data in tracker URL", "high"),
    (r"[?&](q|search|query|keyword)=([^&]{3,})",
     "Search / query string exposed in tracker URL", "medium"),
]

# Health-related path terms that raise payload severity
HEALTH_PATH_TERMS = [
    "appointment", "portal", "patient", "doctor", "physician", "provider",
    "condition", "symptom", "treatment", "diagnosis", "medication",
    "telehealth", "virtual", "mental", "health", "clinic", "hospital",
    "therapy", "intake", "registration", "schedule", "booking",
    "lab", "test", "result", "record", "insurance", "billing",
    "referral", "specialist", "prescription",
]


# ── Colour palettes ───────────────────────────────────────────────────────────
TIMING_COLOURS = {
    "NO-CONSENT-MECHANISM": {"bg": "8B0000", "fg": "FFFFFF"},
    "PRE-BANNER":      {"bg": "FF4444", "fg": "8B0000"},
    "BANNER-PRESENT":  {"bg": "FFD966", "fg": "7D5A00"},
    "NO TRACKER":      {"bg": None,     "fg": "444444"},
}

def risk_bg(score: int) -> str:
    """Return a hex fill colour for a 1-10 risk score."""
    if score >= 9: return "FF4444"   # Critical — red
    if score >= 7: return "FF8C00"   # High — orange
    if score >= 5: return "FFD966"   # Medium — yellow
    return "92D050"                   # Low — green

def risk_label(score: int) -> str:
    if score >= 9: return "CRITICAL"
    if score >= 7: return "HIGH"
    if score >= 5: return "MEDIUM"
    return "LOW"


# ── Timing helper ─────────────────────────────────────────────────────────────
def _now_ms():
    import time
    return time.time() * 1_000


# ══════════════════════════════════════════════════════════════════════════════
# DOMAIN / URL PARSING
# ══════════════════════════════════════════════════════════════════════════════

DOMAIN_RE = re.compile(r"([a-z0-9-]+\.)+[a-z]{2,}", re.I)

FREE_DOMAINS = {
    "gmail.com", "yahoo.com", "aol.com", "hotmail.com", "outlook.com",
    "icloud.com", "protonmail.com", "zoho.com", "yandex.com", "mail.com",
    # Google infrastructure — never a client site
    "google.com", "goo.gl", "maps.app.goo.gl",
}

# Matches every common Google Maps URL format
GOOGLE_MAPS_RE = re.compile(
    r"(?:"
    r"(?:https?://)?(?:www\.)?google\.[a-z]{2,6}(?:\.[a-z]{2})?/maps"
    r"|(?:https?://)?maps\.google\.[a-z]{2,6}(?:\.[a-z]{2})?"
    r"|(?:https?://)?goo\.gl/maps"
    r"|(?:https?://)?maps\.app\.goo\.gl"
    r")",
    re.I,
)

# Matches Google Ads click-through URLs — these REDIRECT to real business sites
GOOGLE_AD_RE = re.compile(
    r"(?:https?://)?(?:www\.)?google\.[a-z]{2,6}(?:\.[a-z]{2})?/aclk",
    re.I,
)


def resolve_redirect(url: str, timeout: int = 10):
    """Follow HTTP redirects and return the final URL, or None on failure."""
    import urllib.request
    try:
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                )
            },
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.url
    except Exception:
        return None


def normalize(raw: str):
    """Extract a clean domain from any raw string.  Returns None for noise."""
    s = re.sub(r"^[a-z]+://", "", raw.strip().strip("\"'"), flags=re.I)
    m = DOMAIN_RE.search(s)
    if not m:
        return None
    d = re.sub(r"^www\.", "", m.group(0).lower().split("/")[0].split(":")[0])
    if "@" in raw:
        d = raw.strip().split("@")[-1].strip().lower()
        m2 = DOMAIN_RE.match(d)
        d = m2.group(0) if m2 else d
    return None if d in FREE_DOMAINS else d


PHONE_RE = re.compile(
    r"(?:\+?\d{1,3}[\s.-]?)?(?:\(\d{2,4}\)[\s.-]?)?\d{3,4}[\s.-]?\d{3,4}[\s.-]?\d{0,4}"
)

def is_phone_line(line: str) -> bool:
    digits = re.sub(r"\D", "", line)
    return bool(re.match(r"^\+?\d[\d\s().-]{6,}$", line)) and 7 <= len(digits) <= 15

def extract_phone(line: str):
    m = PHONE_RE.search(line)
    return m.group(0).strip() if m else None


def extract_entries(text: str) -> list:
    """
    Parse the raw DOMAINS paste and return [{"domain": str, "phone": str}, …].
    Handles bare domains, full URLs, Google Maps links, Google Ads /aclk links,
    email addresses, phone numbers, and general Google-Maps-export noise.
    """
    seen, result = set(), []
    pending_phone = None

    for line in [l.strip() for l in text.splitlines()]:
        if not line:
            continue

        if is_phone_line(line):
            pending_phone = extract_phone(line)
            continue

        # ── Google Ads redirect: follow to get the real business domain ───────
        if GOOGLE_AD_RE.search(line):
            short = line[:70] + "..." if len(line) > 70 else line
            print(f"  [AD LINK] Resolving {short}")
            final = resolve_redirect(line.strip())
            if final:
                d = normalize(final)
                if d and d not in seen:
                    seen.add(d)
                    result.append({"domain": d, "phone": pending_phone or ""})
                    pending_phone = None
                    print(f"  [AD LINK] -> {d}")
            continue

        # ── Google Maps: skip entirely ────────────────────────────────────────
        if GOOGLE_MAPS_RE.search(line):
            continue

        # ── Noise filters ─────────────────────────────────────────────────────
        if re.match(r"^\d+(\.\d+)?\s*(stars?|Reviews?|\(\d+\))", line, re.I):
            continue
        if re.match(r"^\d+\s+\w+", line) and any(
            tok in line for tok in ("St", "Ave", "Rd", "Blvd")
        ):
            continue
        if line.lower() in ("open", "closed") or line.lower().startswith(
            ("closes", "opens")
        ):
            continue

        d = normalize(line)
        if d and d not in seen:
            seen.add(d)
            result.append({"domain": d, "phone": pending_phone or ""})
            pending_phone = None

    return result


# ══════════════════════════════════════════════════════════════════════════════
# PAGE CLASSIFICATION
# ══════════════════════════════════════════════════════════════════════════════

def registrable_domain(netloc: str) -> str:
    """
    Reduce a netloc to its registrable domain for redirect comparison.
    e.g. "www.rokitbenefits.com" -> "rokitbenefits.com"
         "app.example.co.uk"    -> "example.co.uk"  (best-effort, no PSL dependency)
    This is intentionally simple (last two labels, or three for common
    two-part TLDs) since we only need to tell "same business domain" apart
    from "redirected to a genuinely different domain" — not full PSL accuracy.
    """
    host = netloc.lower().split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    labels = host.split(".")
    if len(labels) <= 2:
        return host
    # Common two-part public suffixes (co.uk, com.au, co.in, etc.) — keep 3 labels
    two_part_tlds = {"co", "com", "org", "net", "gov", "ac", "edu"}
    if labels[-2] in two_part_tlds and len(labels[-1]) == 2:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


def classify_page_type(url: str) -> tuple:
    """
    Returns (type_key, label, base_risk_score) for a given URL.
    """
    path = urlparse(url).path.lower()

    # Explicit homepage detection
    if re.match(r"^/?$", path) or path in ("/home", "/home/", "/index.html"):
        return "homepage", "Homepage", PAGE_TYPES["homepage"]["risk"]

    for key, data in PAGE_TYPES.items():
        if key == "homepage":
            continue
        for pat in data["patterns"]:
            if re.search(pat, path, re.I):
                return key, data["label"], data["risk"]

    return "unknown", "General Public Page", 2


def classify_tracker(req_url: str):
    """Return vendor name if the request URL matches a known ad-tech domain."""
    low = req_url.lower()
    for vendor, patterns in TRACKERS.items():
        if any(p in low for p in patterns):
            return vendor
    return None


# ══════════════════════════════════════════════════════════════════════════════
# PAYLOAD SENSITIVITY ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════

def analyze_payload(tracker_url: str, page_url: str) -> list:
    """
    Inspect a tracker request URL for health-related signals.
    Returns a deduplicated list of {"label": str, "severity": "high"|"medium"|"info"} dicts.
    """
    flags = []
    decoded = unquote(tracker_url.lower())

    for pattern, label, severity in SENSITIVE_PATTERNS:
        if re.search(pattern, decoded, re.I):
            flags.append({"label": label, "severity": severity})

    # ── Explicit event-name reporting ───────────────────────────────────────
    # FIX: previously the report only ever spoke up when a HIGH-risk
    # conversion event fired (Lead/Schedule/etc.) and said nothing otherwise
    # — meaning a plain PageView and a withheld Lead inference looked
    # identical in the report (silence). Per review: state the actual
    # observed event by name so "PageView" is reported as PageView, not
    # left to be inferred as something higher-risk by omission.
    ev_match = re.search(r"[?&]ev=([a-z0-9_]+)", decoded, re.I)
    if ev_match:
        observed_event = ev_match.group(1)
        HIGH_RISK_EVENTS = {
            "lead", "completeregistration", "schedule", "contact",
            "initiatecheckout", "purchase", "submitapplication",
            "subscribe", "starttrial", "addpaymentinfo",
        }
        if observed_event.lower() not in HIGH_RISK_EVENTS:
            flags.append({
                "label": f"Event observed: {observed_event} (standard browse/analytics "
                         f"event — not a Lead/Purchase/Schedule/Registration conversion)",
                "severity": "info",
            })
        # (high-risk events are already captured by the anchored ev= pattern
        # in SENSITIVE_PATTERNS above with severity "high" — not duplicated here)

    # Check if the *page* URL's own path contains health terms
    page_path = urlparse(page_url).path.lower()
    found_terms = [t for t in HEALTH_PATH_TERMS if t in page_path]
    if found_terms:
        flags.append({
            "label": f"Page URL exposes health context: {', '.join(found_terms[:4])}",
            "severity": "info",
        })

    # Deduplicate by label while preserving order
    seen_f, unique = set(), []
    for f in flags:
        if f["label"] not in seen_f:
            seen_f.add(f["label"])
            unique.append(f)
    return unique


# ══════════════════════════════════════════════════════════════════════════════
# RISK SCORING
# ══════════════════════════════════════════════════════════════════════════════

def score_risk(page_risk: int, timing: str, payload_flags: list) -> tuple:
    """
    Combine page-type base risk, timing classification, and payload flags
    into a single 1-10 score, PLUS a plain-English breakdown string so the
    score is auditable without reverse-engineering the math.

    Timing penalties:
      NO-CONSENT-MECHANISM → +3  (no consent UI/CMP exists on the page at
                                   all — there was never a chance to consent,
                                   not just "tracker won a race against a
                                   banner". Distinct from, and scored higher
                                   than, PRE-BANNER per privacy-exposure
                                   reasoning: an absent consent flow is a
                                   stronger finding than a present-but-slow one.)
      PRE-BANNER            → +2  (a consent banner/CMP DOES exist on this
                                    page, but this tracker fired before it
                                    rendered)
      BANNER-PRESENT        → +1  (tracker fires despite banner already
                                    being visible)

    Payload penalties (severity-weighted, not just a flag count):
      any "high"-severity flag (real PII / form data / conversion event) → +3
      else any "medium"-severity flag (browser/ad ID only, no real PII)   → +1
      "info"-only flags (health keywords, URL exposure)                  → +0
        (info flags add context but shouldn't inflate the numeric score)

    Returns: (score:int, breakdown:str)
    """
    parts = [f"Base({page_risk})"]
    score = page_risk

    if timing == "NO-CONSENT-MECHANISM":
        score += 3
        parts.append("NO-CONSENT-MECHANISM(+3)")
    elif timing == "PRE-BANNER":
        score += 2
        parts.append("PRE-BANNER(+2)")
    elif timing == "BANNER-PRESENT":
        score += 1
        parts.append("BANNER-PRESENT(+1)")

    severities = {f["severity"] for f in payload_flags}
    if "high" in severities:
        score += 3
        parts.append("Likely-PII/conversion(+3)")
    elif "medium" in severities:
        score += 1
        parts.append("Browser-ID-only(+1)")

    score = max(1, min(10, score))
    parts.append(f"= {score}")
    return score, "  ".join(parts)


# ══════════════════════════════════════════════════════════════════════════════
# PAGE DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════

def discover_pages(domain: str) -> list:
    """
    1. Load homepage in Playwright; harvest all internal <a href> links.
    2. Classify each link; keep up to 2 per page type.
    3. For page types still missing, probe common healthcare URL paths via HTTP.

    FIX: previously, if the homepage redirected to a different registrable
    domain (e.g. rokittelemedicine.com -> rokitbenefits.com), pg.goto()
    silently followed it but base_url / pages[0]["url"] kept the ORIGINAL
    domain. Every subsequent finding then got mislabeled under the wrong
    domain. Now every page-info dict carries:
      - "requested_domain": the domain the user actually typed in
      - "redirected": True/False — whether the final loaded URL's
                       registrable domain differs from the requested one
      - "final_domain": the registrable domain actually loaded, if redirected

    Returns a list of page-info dicts:
      {"url", "page_type", "page_label", "page_risk",
       "requested_domain", "redirected", "final_domain"}
    """
    import urllib.request as _ur

    base_url = f"https://{domain}"
    requested_reg_domain = registrable_domain(domain)
    redirected   = False
    final_domain = requested_reg_domain

    pages = [{
        "url":              base_url,
        "page_type":        "homepage",
        "page_label":       "Homepage",
        "page_risk":        PAGE_TYPES["homepage"]["risk"],
        "requested_domain": requested_reg_domain,
        "redirected":       False,
        "final_domain":     requested_reg_domain,
    }]
    found_types = {"homepage"}

    # ── Step 1: crawl homepage ─────────────────────────────────────────────
    print(f"    Crawling homepage for internal links ...")
    internal_links = []
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            ctx = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                )
            )
            pg = ctx.new_page()
            try:
                pg.goto(base_url, wait_until="domcontentloaded", timeout=20_000)
            except Exception:
                base_url = f"http://{domain}"
                try:
                    pg.goto(base_url, wait_until="domcontentloaded", timeout=15_000)
                except Exception:
                    browser.close()
                    return pages

            # ── FIX: check where the browser actually ended up ─────────────
            landed_url    = pg.url
            landed_netloc = urlparse(landed_url).netloc
            landed_reg    = registrable_domain(landed_netloc)
            if landed_reg != requested_reg_domain:
                redirected   = True
                final_domain = landed_reg
                print(
                    f"    [REDIRECT DETECTED] {domain} -> {landed_netloc} "
                    f"(registrable domain changed: {requested_reg_domain} -> {landed_reg})"
                )

            # Record the ACTUAL loaded URL, plus redirect metadata, on homepage
            pages[0]["url"]          = landed_url
            pages[0]["redirected"]   = redirected
            pages[0]["final_domain"] = final_domain
            base_url = landed_url   # subsequent probes should target where we actually landed

            raw_links = pg.eval_on_selector_all(
                "a[href]", "els => els.map(e => e.href)"
            )
            browser.close()

        # Internal-link filter now compares against the domain we actually
        # landed on (final_domain), not the originally requested one —
        # otherwise every link on a redirected site would be (wrongly)
        # treated as "external" and discovery would find nothing.
        for link in raw_links:
            try:
                p_link = urlparse(link)
                link_reg = registrable_domain(p_link.netloc)
                if link_reg == final_domain and p_link.scheme in ("http", "https"):
                    internal_links.append(link)
            except Exception:
                pass

        # Classify and add up to 2 per type
        type_counts = {}
        for link in set(internal_links):
            ptype, plabel, prisk = classify_page_type(link)
            if ptype in ("unknown", "homepage"):
                continue
            if type_counts.get(ptype, 0) < 2:
                pages.append({
                    "url":              link,
                    "page_type":        ptype,
                    "page_label":       plabel,
                    "page_risk":        prisk,
                    "requested_domain": requested_reg_domain,
                    "redirected":       redirected,
                    "final_domain":     final_domain,
                })
                type_counts[ptype] = type_counts.get(ptype, 0) + 1
                found_types.add(ptype)

    except Exception as e:
        print(f"    [DISCOVERY ERROR] {e}")

    print(
        f"    {len(pages)-1} non-homepage page(s) found via crawl. "
        f"Probing for missing types ..."
    )

    # ── Step 2: probe missing page types ──────────────────────────────────
    # Probes target base_url, which now points at wherever we actually
    # landed (post-redirect) so probed paths resolve against the real site.
    for ptype, paths in PROBE_PATHS.items():
        if ptype in found_types:
            continue
        for path in paths:
            probe_url = base_url.rstrip("/") + path
            try:
                req = _ur.Request(
                    probe_url,
                    headers={"User-Agent": "Mozilla/5.0", "Accept": "text/html"},
                )
                with _ur.urlopen(req, timeout=6) as resp:
                    final_probe_url = resp.url   # follow any further redirect on the probe itself
                    if resp.status == 200:
                        pdata = PAGE_TYPES[ptype]
                        pages.append({
                            "url":              final_probe_url,
                            "page_type":        ptype,
                            "page_label":       pdata["label"],
                            "page_risk":        pdata["risk"],
                            "requested_domain": requested_reg_domain,
                            "redirected":       redirected,
                            "final_domain":     final_domain,
                        })
                        found_types.add(ptype)
                        print(f"    [PROBE HIT] {ptype} -> {final_probe_url}")
                        break
            except Exception:
                continue

    return pages


# ══════════════════════════════════════════════════════════════════════════════
# SINGLE-PAGE SCANNER
# ══════════════════════════════════════════════════════════════════════════════

def scan_page(url: str, page_type: str, page_label: str, page_risk: int) -> dict:
    """
    Open *url* as a first-time visitor in a completely fresh browser context
    (no cookies, no session state).  Do NOT click the consent banner.

    Records:
      - All outbound tracker requests + timing relative to page load
      - Whether a consent banner is visible before each tracker fires
      - All cookies set during the visit (flagging known ad-tech cookies)

    Returns a structured result dict consumed by build_excel().
    """
    tracker_hits      = []
    total_requests    = 0
    consent_at_ms     = None      # timestamp when banner text detected
    consent_ever_appeared = False # True only if a consent mechanism appeared at ANY point in the visit
    start_ts          = None
    ad_cookies_found  = []
    error             = None
    landed_url        = url        # updated to the actual URL after navigation
    page_redirected   = False      # True if landed_url's registrable domain != requested url's

    # Known ad-tech cookie name fragments
    AD_COOKIE_SIGNALS = (
        "_fbp", "_ga", "_gid", "fbp", "fbc", "ttclid", "_ttp",
        "_uetsid", "_hjid", "_gcl_au", "ba_vid",
    )

    def on_request(request):
        nonlocal total_requests
        total_requests += 1
        vendor = classify_tracker(request.url)
        if vendor:
            elapsed_ms = int(_now_ms() - start_ts) if start_ts else 0
            pre_banner  = (consent_at_ms is None)
            tracker_hits.append({
                "vendor":     vendor,
                "url":        request.url,
                "elapsed_ms": elapsed_ms,
                "pre_banner": pre_banner,
            })

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                viewport   = {"width": 1280, "height": 800},
                user_agent = (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
            )
            pg = context.new_page()
            pg.on("request", on_request)

            start_ts = _now_ms()
            try:
                pg.goto(url, wait_until="domcontentloaded", timeout=25_000)
                # ── FIX: detect if THIS individual page redirected somewhere
                # with a different registrable domain than what was requested.
                # Catches cases discover_pages() can't: a probed/crawled URL
                # that itself 301s elsewhere independent of the homepage.
                landed_url = pg.url
                req_reg    = registrable_domain(urlparse(url).netloc)
                landed_reg = registrable_domain(urlparse(landed_url).netloc)
                if landed_reg != req_reg:
                    page_redirected = True
                    print(
                        f"         [PAGE REDIRECT] {url} -> {landed_url} "
                        f"({req_reg} -> {landed_reg})"
                    )
            except Exception as e:
                error = str(e)

            def _consent_mechanism_visible() -> bool:
                """
                Check for a consent mechanism via TWO independent signals:
                  1. CONSENT_KEYWORDS in visible body text (catches custom/
                     homegrown banners with no recognizable script tag).
                  2. CMP_SCRIPT_SIGNATURES in the page's <script src> tags
                     or inline script content (catches OneTrust/Cookiebot/
                     CookieYes/etc. even if their banner text hasn't
                     rendered yet, or uses non-English / unusual wording).
                Returns True if EITHER signal fires.
                """
                try:
                    body = pg.inner_text("body").lower()
                    if any(k in body for k in CONSENT_KEYWORDS):
                        return True
                except Exception:
                    pass
                try:
                    script_srcs = pg.eval_on_selector_all(
                        "script", "els => els.map(e => e.src || '')"
                    )
                    combined = " ".join(script_srcs).lower()
                    if any(sig in combined for sig in CMP_SCRIPT_SIGNATURES):
                        return True
                except Exception:
                    pass
                return False

            # Detect consent banner/CMP presence — read only, no click.
            # FIX: banners that appear after a delay (2-10s) or on scroll were
            # being missed because the body was checked immediately on load.
            # Wait up to 10s, polling, before concluding no banner exists.
            try:
                for _ in range(10):
                    if _consent_mechanism_visible():
                        consent_at_ms = _now_ms()
                        break
                    pg.wait_for_timeout(1_000)
            except Exception:
                pass

            # Observe deferred / lazy pixel fires for the remainder of a
            # 15-second total observation window (10s already spent above).
            pg.wait_for_timeout(5_000)

            # Final check: some banners load even later via scroll-trigger
            # or delayed script injection. Re-check once more before closing.
            # FIX: this is also the authoritative "did a consent mechanism
            # ever appear during the ENTIRE observation window" check, used
            # below to distinguish PRE-BANNER from NO-CONSENT-MECHANISM.
            if consent_at_ms is None:
                if _consent_mechanism_visible():
                    consent_at_ms = _now_ms()
            consent_ever_appeared = consent_at_ms is not None

            # Collect ad-tech cookies planted during this visit
            try:
                for c in context.cookies():
                    if any(s in c["name"].lower() for s in AD_COOKIE_SIGNALS):
                        ad_cookies_found.append(c["name"])
            except Exception:
                pass

            browser.close()

    except Exception as e:
        error = str(e)

    # ── Build per-finding records ──────────────────────────────────────────
    # FIX: previously "PRE-BANNER" meant only "no banner had appeared YET at
    # fire-time" — which is the same label whether a banner later showed up
    # (tracker just beat it) or no consent mechanism existed on the page at
    # all, ever. Those are different findings with different implications.
    # Now: if consent_ever_appeared is False (no banner/CMP detected during
    # the ENTIRE observation window), every tracker hit on this page is
    # reclassified as NO-CONSENT-MECHANISM rather than PRE-BANNER, even
    # though it was provisionally tagged "pre_banner" at fire-time.
    findings = []
    vendors_seen = set()

    for hit in tracker_hits:
        if not consent_ever_appeared:
            timing = "NO-CONSENT-MECHANISM"
        else:
            timing = "PRE-BANNER" if hit["pre_banner"] else "BANNER-PRESENT"
        # Use landed_url (not the originally-requested url) for page-path
        # context — if the page redirected, the health-context keywords
        # should reflect the page that ACTUALLY loaded.
        pflags  = analyze_payload(hit["url"], landed_url)
        risk, breakdown = score_risk(page_risk, timing, pflags)
        vendors_seen.add(hit["vendor"])
        findings.append({
            "vendor":          hit["vendor"],
            "request_url":     hit["url"][:300],
            "elapsed_ms":      hit["elapsed_ms"],
            "timing":          timing,
            "payload_flags":   pflags,
            "risk_score":      risk,
            "risk_label":      risk_label(risk),
            "risk_breakdown":  breakdown,
        })

    # ── Page-level timing summary ──────────────────────────────────────────
    # Priority order when multiple timings occur across hits on one page:
    # NO-CONSENT-MECHANISM is reported whenever it's present (most severe —
    # means literally no consent flow exists), then PRE-BANNER (banner
    # exists but was beaten), then BANNER-PRESENT.
    if not findings:
        timing_summary = "NO TRACKER"
        max_risk = page_risk  # page still carries inherent risk
        max_risk_breakdown = f"Base({page_risk})  No tracker observed  = {page_risk}"
    else:
        if any(f["timing"] == "NO-CONSENT-MECHANISM" for f in findings):
            timing_summary = "NO-CONSENT-MECHANISM"
        elif any(f["timing"] == "PRE-BANNER" for f in findings):
            timing_summary = "PRE-BANNER"
        else:
            timing_summary = "BANNER-PRESENT"
        worst_finding = max(findings, key=lambda f: f["risk_score"])
        max_risk = worst_finding["risk_score"]
        max_risk_breakdown = worst_finding["risk_breakdown"]

    return {
        "url":               landed_url,    # the URL actually scanned/loaded
        "requested_url":     url,           # the URL we asked the browser to open
        "page_redirected":   page_redirected,
        "page_type":         page_type,
        "page_label":        page_label,
        "page_risk":         page_risk,
        "consent_banner":    consent_at_ms is not None,
        "timing_summary":    timing_summary,
        "vendors":           sorted(vendors_seen),
        "findings":          findings,
        "ad_cookies":        ad_cookies_found,
        "max_risk_score":    max_risk,
        "max_risk_label":    risk_label(max_risk),
        "max_risk_breakdown": max_risk_breakdown,
        "total_requests":   total_requests,
        "error":             error,
    }


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL REPORT BUILDER
# ══════════════════════════════════════════════════════════════════════════════

def _thin_border(colour="CCCCCC"):
    s = Side(style="thin", color=colour)
    return Border(left=s, right=s, top=s, bottom=s)

def _c(ws, row, col, value,
       bg=None, fg="111111", bold=False,
       align="left", wrap=False, size=10, border=True):
    """Write a styled cell."""
    cell = ws.cell(row=row, column=col, value=value)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    cell.font      = Font(name="Arial", size=size, bold=bold, color=fg)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if border:
        cell.border = _thin_border()
    return cell


def build_excel(domain_results: list, path: str, category: str = ""):
    """
    Build a three-sheet Excel workbook.

    Sheet 1 — Executive Summary   (one row per domain)
    Sheet 2 — Page Findings       (one row per scanned page)
    Sheet 3 — Tracker Detail      (one row per individual tracker hit)
    """
    wb  = Workbook()
    HDR = "1A1A2E"   # dark navy header fill
    BNR = "2E3A59"   # medium navy banner fill

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M")

    # ══ Sheet 1: Executive Summary ════════════════════════════════════════════
    ws1       = wb.active
    ws1.title = "Executive Summary"

    ws1.merge_cells("A1:N1")
    c = ws1["A1"]
    c.value     = (f"Healthcare Pixel Privacy Audit   |   "
                   f"{category or 'Healthcare'}   |   Run: {run_time}")
    c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=13)
    c.fill      = PatternFill("solid", fgColor=BNR)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 30

    # Sub-header note row
    ws1.merge_cells("A2:N2")
    c2 = ws1["A2"]
    c2.value     = ("Risk scale: 10=CRITICAL (patient portal + no consent mechanism / pre-banner tracker)  "
                    "→  1=LOW (public marketing page, no tracker).  "
                    "Base risk = page type; +3 if NO-CONSENT-MECHANISM (no banner/CMP ever detected); "
                    "+2 if PRE-BANNER (banner exists, tracker fired first); +1 if BANNER-PRESENT; "
                    "+1-3 for payload sensitivity flags.")
    c2.font      = Font(italic=True, color="FFFFFF", name="Arial", size=9)
    c2.fill      = PatternFill("solid", fgColor="3D4F73")
    c2.alignment = Alignment(horizontal="left", vertical="center")
    ws1.row_dimensions[2].height = 18

    s1_headers = ["Domain\n(Requested)", "Domain\nRedirected?", "Actual Domain\nScanned",
                  "Mobile\nNumber", "Pages\nScanned", "No-Consent-\nMechanism Pages",
                  "Pre-Banner\nPages",
                  "Banner-Present\nPages", "No-Tracker\nPages",
                  "Vendors Detected", "Highest-Risk\nPage Type",
                  "Max Risk\nScore (1-10)", "Risk Level",
                  "Consent Banner\nDetected"]
    for col, h in enumerate(s1_headers, 1):
        _c(ws1, 3, col, h, bg=HDR, fg="FFFFFF", bold=True,
           align="center", size=10, wrap=True)
    ws1.row_dimensions[3].height = 38

    row = 4
    for dr in domain_results:
        pages = dr["pages"]
        no_consent = [p for p in pages if p["timing_summary"] == "NO-CONSENT-MECHANISM"]
        pre   = [p for p in pages if p["timing_summary"] == "PRE-BANNER"]
        pres  = [p for p in pages if p["timing_summary"] == "BANNER-PRESENT"]
        none_ = [p for p in pages if p["timing_summary"] == "NO TRACKER"]
        vendors  = sorted({v for p in pages for v in p["vendors"]})
        max_risk = max((p["max_risk_score"] for p in pages), default=0)
        worst    = max(pages, key=lambda p: p["max_risk_score"], default=None)
        any_banner = any(p["consent_banner"] for p in pages)
        bg = risk_bg(max_risk)

        any_redirected = any(p.get("redirected") for p in pages)
        # All redirected pages on a domain should point to the same final
        # domain (it's the same site); take the first one found as the
        # representative "actual domain scanned" value for this row.
        final_domains = sorted({
            p.get("final_domain") for p in pages
            if p.get("redirected") and p.get("final_domain")
        })
        actual_domain_display = ", ".join(final_domains) if final_domains else dr["domain"]

        _c(ws1, row, 1,  dr["domain"], bold=True)
        _c(ws1, row, 2,  "YES" if any_redirected else "NO",
           bg="FFD966" if any_redirected else None,
           fg="7D5A00" if any_redirected else "006600",
           bold=True, align="center")
        _c(ws1, row, 3,  actual_domain_display,
           bold=any_redirected, align="center", wrap=True, size=9)
        _c(ws1, row, 4,  dr.get("phone") or "—", align="center")
        _c(ws1, row, 5,  len(pages),   align="center")
        _c(ws1, row, 6,  len(no_consent),
           bg="8B0000" if no_consent else None,
           fg="FFFFFF" if no_consent else "111111",
           bold=bool(no_consent), align="center")
        _c(ws1, row, 7,  len(pre),
           bg="FF4444" if pre else None,
           fg="FFFFFF"  if pre else "111111",
           bold=bool(pre), align="center")
        _c(ws1, row, 8,  len(pres),   align="center")
        _c(ws1, row, 9,  len(none_),  align="center")
        _c(ws1, row, 10, ", ".join(vendors) or "None", wrap=True, size=9)
        _c(ws1, row, 11, worst["page_label"] if worst else "—", align="center")
        _c(ws1, row, 12, max_risk, bg=bg, fg="000000", bold=True, align="center")
        _c(ws1, row, 13, risk_label(max_risk),
           bg=bg, fg="000000", bold=True, align="center")
        _c(ws1, row, 14, "YES" if any_banner else "NO",
           fg="CC0000" if any_banner else "006600",
           bold=True, align="center")

        ws1.row_dimensions[row].height = 22
        row += 1

    col_widths1 = [26, 13, 22, 16, 9, 13, 10, 13, 10, 48, 26, 14, 12, 14]
    for i, w in enumerate(col_widths1, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w
    ws1.freeze_panes = ws1.cell(row=4, column=1)

    # ══ Sheet 2: Page Findings ════════════════════════════════════════════════
    ws2       = wb.create_sheet("Page Findings")
    ws2.merge_cells("A1:O1")
    c = ws2["A1"]
    c.value     = ("Page-by-Page Findings  —  each row = one scanned page  —  "
                   "sorted as discovered (highest-risk page types listed first by design)")
    c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    c.fill      = PatternFill("solid", fgColor=BNR)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 24

    s2_headers = ["Domain\n(Requested)", "Redirected?\n(Domain Changed)", "Page URL", "Page Type", "Base\nRisk",
                  "Consent\nBanner", "Timing\nClassification",
                  "Tracker Vendors\nFound", "Ad Cookies\nPlanted",
                  "Likely PII\nFound?", "Browser/Ad-ID\nFlags (count)",
                  "Payload Flag Summary",
                  "Risk\nScore", "Risk\nLevel", "Risk Score\nBreakdown"]
    for col, h in enumerate(s2_headers, 1):
        _c(ws2, 3, col, h, bg=HDR, fg="FFFFFF", bold=True,
           align="center", size=10, wrap=True)
    ws2.row_dimensions[3].height = 36

    row = 4
    ctr = 1
    for dr in domain_results:
        for page in dr["pages"]:
            timing = page["timing_summary"]
            tc     = TIMING_COLOURS.get(timing, {"bg": None, "fg": "111111"})
            base_row_bg = tc["bg"] or ("F5F5F5" if ctr % 2 == 0 else "FFFFFF")
            rs = page["max_risk_score"]

            all_flags = list({
                flag["label"]
                for f in page["findings"]
                for flag in f["payload_flags"]
            })
            has_high_pii = any(
                flag["severity"] == "high"
                for f in page["findings"] for flag in f["payload_flags"]
            )
            medium_count = sum(
                1 for f in page["findings"] for flag in f["payload_flags"]
                if flag["severity"] == "medium"
            )
            flag_summary = " | ".join(all_flags[:3]) + (
                f" (+{len(all_flags)-3} more)" if len(all_flags) > 3 else ""
            ) if all_flags else "—"

            page_redirected = page.get("redirected", False)

            _c(ws2, row, 1,  dr["domain"])
            _c(ws2, row, 2,  "YES" if page_redirected else "NO",
               bg="FFD966" if page_redirected else None,
               fg="7D5A00" if page_redirected else "006600",
               bold=True, align="center")
            _c(ws2, row, 3,  page["url"], wrap=True, size=9,
               fg="7D5A00" if page_redirected else "111111",
               bold=page_redirected)
            _c(ws2, row, 4,  page["page_label"], align="center")
            _c(ws2, row, 5,  page["page_risk"],  align="center")
            _c(ws2, row, 6,  "YES" if page["consent_banner"] else "NO",
               fg="CC0000" if page["consent_banner"] else "006600",
               bold=True, align="center")
            _c(ws2, row, 7,  timing,
               bg=tc["bg"], fg=tc["fg"] or "111111",
               bold=True, align="center")
            _c(ws2, row, 8,  ", ".join(page["vendors"]) or "—",
               wrap=True, size=9)
            _c(ws2, row, 9,  len(page["ad_cookies"]),
               bg="FF4444" if page["ad_cookies"] else None,
               fg="FFFFFF"  if page["ad_cookies"] else "111111",
               align="center")
            _c(ws2, row, 10, "YES" if has_high_pii else "NO",
               bg="FF4444" if has_high_pii else None,
               fg="FFFFFF" if has_high_pii else "006600",
               bold=True, align="center")
            _c(ws2, row, 11, medium_count,
               fg="7D5A00" if medium_count else "111111",
               bold=medium_count > 0, align="center")
            _c(ws2, row, 12, flag_summary, wrap=True, size=9)
            _c(ws2, row, 13, rs,
               bg=risk_bg(rs), fg="000000", bold=True, align="center")
            _c(ws2, row, 14, risk_label(rs),
               bg=risk_bg(rs), fg="000000", bold=True, align="center")
            worst_f = max(page["findings"], key=lambda f: f["risk_score"], default=None)
            _c(ws2, row, 15,
               worst_f["risk_breakdown"] if worst_f else page.get("max_risk_breakdown", "—"),
               wrap=True, size=8)

            ws2.row_dimensions[row].height = 32
            row += 1
            ctr += 1

    col_widths2 = [22, 14, 36, 20, 8, 9, 17, 30, 10, 10, 12, 46, 9, 11, 34]
    for i, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = w
    ws2.freeze_panes = ws2.cell(row=4, column=1)

    # ══ Sheet 3: Tracker Detail ═══════════════════════════════════════════════
    ws3       = wb.create_sheet("Tracker Detail")
    ws3.merge_cells("A1:K1")
    c = ws3["A1"]
    c.value     = ("Individual Tracker Hits  —  each row = one network request to an "
                   "ad-tech vendor  —  check 'Payload Sensitivity Flags' for HIPAA/FTC risk signals")
    c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    c.fill      = PatternFill("solid", fgColor=BNR)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 24

    s3_headers = ["Domain", "Page URL", "Page Type", "Vendor",
                  "Timing\nClassification", "Elapsed\n(ms)",
                  "Tracker Request URL (first 280 chars)",
                  "Payload Sensitivity Flags",
                  "Risk\nScore", "Risk\nLevel", "Risk Score\nBreakdown"]
    for col, h in enumerate(s3_headers, 1):
        _c(ws3, 3, col, h, bg=HDR, fg="FFFFFF", bold=True,
           align="center", size=10, wrap=True)
    ws3.row_dimensions[3].height = 36

    row = 4
    ctr = 1
    for dr in domain_results:
        for page in dr["pages"]:
            if not page["findings"]:
                continue
            for f in page["findings"]:
                timing = f["timing"]
                tc     = TIMING_COLOURS.get(timing, {"bg": None, "fg": "111111"})
                rs     = f["risk_score"]

                flag_labels = [fl["label"] for fl in f["payload_flags"]]
                has_high    = any(fl["severity"] == "high" for fl in f["payload_flags"])
                has_medium  = any(fl["severity"] == "medium" for fl in f["payload_flags"])

                _c(ws3, row, 1,  dr["domain"])
                _c(ws3, row, 2,  page["url"], wrap=True, size=9)
                _c(ws3, row, 3,  page["page_label"], align="center")
                _c(ws3, row, 4,  f["vendor"], bold=True)
                _c(ws3, row, 5,  timing,
                   bg=tc["bg"], fg=tc["fg"] or "111111",
                   bold=True, align="center")
                _c(ws3, row, 6,  f["elapsed_ms"], align="center")
                _c(ws3, row, 7,  f["request_url"][:280], wrap=True, size=8)
                _c(ws3, row, 8,
                   " | ".join(flag_labels) if flag_labels else "—",
                   wrap=True, size=9,
                   fg="CC0000" if has_high else ("7D5A00" if has_medium else "444444"))
                _c(ws3, row, 9,  rs,
                   bg=risk_bg(rs), fg="000000", bold=True, align="center")
                _c(ws3, row, 10, f["risk_label"],
                   bg=risk_bg(rs), fg="000000", bold=True, align="center")
                _c(ws3, row, 11, f["risk_breakdown"], wrap=True, size=8)

                ws3.row_dimensions[row].height = 36
                row += 1
                ctr += 1

    col_widths3 = [26, 30, 20, 22, 16, 8, 50, 50, 9, 11, 34]
    for i, w in enumerate(col_widths3, 1):
        ws3.column_dimensions[get_column_letter(i)].width = w
    ws3.freeze_panes = ws3.cell(row=4, column=1)

    wb.save(path)
    print(f"\n  Saved -> {path}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    category = BUSINESS_TYPE.strip()
    city     = CITY.strip()

    entries = extract_entries(DOMAINS)
    if not entries:
        print(
            "\nNo valid domains found in DOMAINS.\n"
            "Paste your list and re-run.\n"
        )
        return

    print("\nHealthcare Pixel Privacy Auditor")
    print("=" * 60)
    print(f"  Domains to audit : {len(entries)}")
    print(f"  Industry         : {category or '(not set)'}")
    print(f"  City             : {city     or '(not set)'}")
    print(f"  Page types scoped: patient portal, scheduling, telehealth,")
    print(f"                     login, intake/registration, contact forms,")
    print(f"                     condition/treatment, provider pages, homepage")
    print("=" * 60 + "\n")

    safe_cat = (
        re.sub(r"[^\w\s-]", "", category).strip().replace(" ", "_")
        or "healthcare"
    )
    out_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        f"pixel_audit_{safe_cat}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    )

    domain_results = []

    for idx, entry in enumerate(entries, 1):
        domain = entry["domain"]
        print(f"\n[{idx}/{len(entries)}] ▶  Auditing: {domain}")
        print("─" * 60)

        # ── 1. Discover pages ──────────────────────────────────────────────
        pages_to_scan = discover_pages(domain)
        print(f"  → {len(pages_to_scan)} page(s) queued for scanning\n")

        # ── 2. Scan each page ──────────────────────────────────────────────
        scanned = []
        requested_domain = registrable_domain(domain)
        for pi, pinfo in enumerate(pages_to_scan, 1):
            url    = pinfo["url"]
            ptype  = pinfo["page_type"]
            plabel = pinfo["page_label"]
            prisk  = pinfo["page_risk"]

            print(f"  [{pi:>2}/{len(pages_to_scan)}] {plabel:<30}  {url}")
            result = scan_page(url, ptype, plabel, prisk)

            # ── Consolidate redirect signal ────────────────────────────────
            # A page can be flagged as redirected two ways: (1) discover_pages()
            # already detected the HOMEPAGE redirected to a different domain
            # (pinfo["redirected"]), which applies to every page found via
            # that crawl/probe; or (2) scan_page() independently found THIS
            # specific URL redirected when it (re)loaded it just now. Either
            # signal means the report must not silently attribute findings
            # to the originally-requested domain.
            discovery_redirected = pinfo.get("redirected", False)
            page_level_redirected = result.get("page_redirected", False)
            result["redirected"] = discovery_redirected or page_level_redirected
            result["requested_domain"] = requested_domain
            result["final_domain"] = (
                pinfo.get("final_domain")
                or registrable_domain(urlparse(result["url"]).netloc)
            )

            if result["redirected"]:
                print(
                    f"         ⚠  [DOMAIN REDIRECT] requested '{requested_domain}' "
                    f"but page served from '{result['final_domain']}'"
                )

            if result["error"] and not result["findings"]:
                print(f"         ⚠  Error: {result['error'][:80]}")

            if result["findings"]:
                for f in result["findings"]:
                    tag = f"[{f['timing']:<14}]"
                    print(
                        f"         {tag} {f['vendor']:<28} "
                        f"risk {f['risk_score']:>2}/10  ({f['risk_label']})"
                    )
                    for flag in f["payload_flags"]:
                        sev_tag = {"high": "PII", "medium": "ID", "info": "ctx"}[flag["severity"]]
                        print(f"             ⚠  [{sev_tag}] {flag['label']}")
            else:
                print(f"         [NO TRACKER    ]")

            if result["ad_cookies"]:
                print(
                    f"         Ad cookies planted: "
                    f"{', '.join(result['ad_cookies'][:4])}"
                )

            scanned.append(result)

        domain_results.append({
            "domain":   domain,
            "phone":    entry.get("phone", ""),
            "category": category,
            "city":     city,
            "pages":    scanned,
        })

    # ── 3. Write Excel report ──────────────────────────────────────────────
    print(f"\n\nBuilding Excel report ...")
    build_excel(domain_results, out_path, category)
    print(f"\nDone!  Open: {os.path.basename(out_path)}\n")


if __name__ == "__main__":
    main()