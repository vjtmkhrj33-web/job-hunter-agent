import html as html_lib
import ipaddress
import json
import os
import re
import socket
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent
TOOLS_DIR = BASE_DIR / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import posting_guard  # noqa: E402
from fetch_page import fetch_with_urllib, strip_tags as fp_strip_tags  # noqa: E402
from job_key import make_key  # noqa: E402
from page_decode import decode_page  # noqa: E402
from rank_state import STATE as SEEN_STATE_PATH, TRACKER as TRACKER_PATH, band, overall_score, save_state, tracker_pairs  # noqa: E402
from robots_check import gate as robots_gate  # noqa: E402
from scrape_state import select as scrape_select  # noqa: E402

LINKEDIN_SEARCH_URL = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
LINKEDIN_DETAIL_URL = "https://www.linkedin.com/jobs-guest/jobs/api/jobPosting"
FREEHIRE_BASE_URL = os.environ.get("FREEHIRE_API_URL", "https://freehire.me").strip().rstrip("/") or "https://freehire.me"

LINKEDIN_UA = "Mozilla/5.0 (compatible; linkedin-search-cli/1.0)"
FREEHIRE_UA = "ai-job-search-dev/ai-job-search/1.0 (freehire-search)"

COMPANY_DISPLAY_MAP = {
    "razorpaysoftwareprivatelimited": "Razorpay",
    "razorpay software private limited": "Razorpay",
    "currenciesdirect": "Redpin (Currencies Direct)",
    "gokwik commerce solutions pvt. ltd.": "GoKwik",
    "converse placement services india private limited": "Converse Placement Services",
    "societe generale global solution centre": "Societe Generale",
}

KNOWN_GREENHOUSE_DOMAIN_BOARDS = {
    "stripe.com": "stripe",
    "careers.godaddy": "godaddy",
    "godaddy.com": "godaddy",
    "careers.airbnb.com": "airbnb",
    "careers.coinbase.com": "coinbase",
    "razorpay.com": "razorpaysoftwareprivatelimited",
}


def normalize_company_display(raw: str) -> str:
    """Turn ATS company slugs (e.g. razorpaysoftwareprivatelimited) into clean company names."""
    s = (raw or "").strip()
    if not s:
        return ""
    low = s.lower()
    if low in COMPANY_DISPLAY_MAP:
        return COMPANY_DISPLAY_MAP[low]
    # Un-slug single lowercase run ending in legal words if present
    if " " not in s and low.endswith("privatelimited"):
        base = re.sub(r"(?:software)?privatelimited$", "", s, flags=re.I).strip()
        if base:
            return base.title()
    return s


def _company_search_term(raw: str) -> str:
    """Clean company name for LinkedIn People Search queries."""
    clean = normalize_company_display(raw)
    clean = re.sub(
        r"\b(?:private\s+limited|pvt\.?\s*ltd\.?|limited|ltd\.?|inc\.?|llc|corp\.?|global\s+solution\s+centre)\b",
        "",
        clean,
        flags=re.I,
    )
    return re.sub(r"\s+", " ", clean).strip(" ,.-") or clean


def is_safe_public_url(url: str) -> tuple[bool, str]:
    """Validate http(s) URL and block loopback/private/link-local hosts (SSRF guard)."""
    try:
        parsed = urllib.parse.urlsplit(url.strip())
    except Exception:
        return False, "Invalid URL syntax"
    if parsed.scheme.lower() not in ("http", "https") or not parsed.netloc:
        return False, "URL must start with http:// or https://"
    hostname = (parsed.hostname or "").strip().lower()
    if not hostname or hostname in ("localhost", "127.0.0.1", "::1", "0.0.0.0"):
        return False, "Localhost/loopback URLs are blocked"
    try:
        ip = ipaddress.ip_address(hostname)
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            return False, "Private or reserved IP targets are blocked"
    except ValueError:
        try:
            addrs = socket.getaddrinfo(hostname, None, proto=socket.IPPROTO_TCP)
            for addr in addrs:
                ip_str = addr[4][0].split("%")[0]
                ip = ipaddress.ip_address(ip_str)
                if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
                    return False, f"Resolved host {hostname} points to a non-public IP"
        except socket.gaierror:
            # Let downstream HTTP request report DNS resolution error if offline
            pass
    return True, ""


def _http_get_with_retry(url: str, headers: dict, max_retries: int = 3, timeout: float = 12.0) -> tuple[int, str]:
    delay = 0.5
    last_err = None
    for attempt in range(max_retries + 1):
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read(5 * 1024 * 1024)
                return getattr(resp, "status", 200), decode_page(raw)
        except urllib.error.HTTPError as exc:
            code = exc.code
            if code == 404:
                return 404, ""
            if code == 429 or code >= 500:
                last_err = f"HTTP {code}: {exc.reason}"
                if attempt < max_retries:
                    time.sleep(min(delay, 3.0))
                    delay *= 2
                    continue
                return code, ""
            return code, decode_page(exc.read(1024 * 256))
        except Exception as exc:
            last_err = str(exc)
            if attempt < max_retries:
                time.sleep(delay)
                delay *= 2
                continue
            raise RuntimeError(last_err)
    raise RuntimeError(last_err or "HTTP request failed")


def _clean_fragment(fragment: str) -> str:
    cleaned, _ = posting_guard.clean_html(fragment)
    cleaned, _ = posting_guard.clean_text(cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def normalize_linkedin_id(raw: str) -> str | None:
    s = (raw or "").strip()
    m = re.search(r"urn:li:jobPosting:(\d+)", s)
    if m:
        return m.group(1)
    if re.fullmatch(r"\d{6,}", s):
        return s
    try:
        parsed = urllib.parse.urlsplit(s if "://" in s else f"https://{s}")
        host = (parsed.hostname or "").lower()
        if host == "linkedin.com" or host.endswith(".linkedin.com"):
            pm = re.search(r"/jobs/view/(?:[^/]*-)?(\d{6,})/?", parsed.path)
            if pm:
                return pm.group(1)
    except Exception:
        pass
    sm = re.search(r"-(\d{6,})$", s)
    return sm.group(1) if sm else None


def normalize_freehire_slug(raw: str) -> str | None:
    s = (raw or "").strip()
    if not s:
        return None
    if re.fullmatch(r"[a-z0-9][a-z0-9-]*", s, re.I):
        return s
    try:
        parsed = urllib.parse.urlsplit(s if "://" in s else f"https://{s}")
        host = (parsed.hostname or "").lower()
        fh_host = (urllib.parse.urlsplit(FREEHIRE_BASE_URL).hostname or "freehire.me").lower()
        if host == "freehire.me" or host.endswith(".freehire.me") or host == fh_host:
            pm = re.match(r"^/jobs/([a-z0-9][a-z0-9-]*)/?$", parsed.path, re.I)
            if pm:
                return pm.group(1)
    except Exception:
        pass
    return None


def parse_linkedin_cards(raw_html: str) -> list[dict]:
    results = []
    chunks = re.split(r'data-entity-urn="urn:li:jobPosting:', raw_html)[1:]
    for chunk in chunks:
        id_m = re.match(r"^(\d+)", chunk)
        if not id_m:
            continue
        job_id = id_m.group(1)
        link_m = re.search(r'class="base-card__full-link[^"]*"[^>]*href="([^"]+)"', chunk, re.I)
        url = html_lib.unescape(link_m.group(1)).split("?")[0] if link_m else f"https://www.linkedin.com/jobs/view/{job_id}"

        title = ""
        h3 = re.search(r'class="base-search-card__title"[^>]*>([\s\S]*?)</h3>', chunk, re.I)
        if h3:
            title = _clean_fragment(h3.group(1))
        if not title:
            sr = re.search(r'class="sr-only"[^>]*>([\s\S]*?)</span>', chunk, re.I)
            if sr:
                title = _clean_fragment(sr.group(1))
        if not title:
            continue

        company = None
        sub = re.search(r'class="base-search-card__subtitle"[^>]*>([\s\S]*?)</h4>', chunk, re.I)
        if sub:
            company = _clean_fragment(sub.group(1)) or None

        loc = re.search(r'class="job-search-card__location"[^>]*>([\s\S]*?)</span>', chunk, re.I)
        location = _clean_fragment(loc.group(1)) if loc else None

        dt = re.search(r'class="job-search-card__listdate[^"]*"[^>]*datetime="([^"]+)"', chunk, re.I)
        posted_date = dt.group(1)[:10] if dt else None

        results.append({
            "id": job_id,
            "title": title,
            "company": company or "Unknown Company",
            "location": location or "",
            "date": posted_date,
            "url": url,
            "portal": "linkedin-search",
            "source": "cli",
            "description": "",
        })
    return results


def search_linkedin(query: str, location: str = "Bengaluru, Karnataka, India", jobage_days: int = 14, remote: str | None = None, page: int = 1, limit: int = 20) -> dict:
    params = {}
    if query:
        params["keywords"] = query
    if location:
        params["location"] = location
    if jobage_days and 0 < jobage_days < 9999:
        params["f_TPR"] = f"r{jobage_days * 86400}"
    wt_map = {"onsite": "1", "on-site": "1", "remote": "2", "hybrid": "3"}
    if remote and remote.lower() in wt_map:
        params["f_WT"] = wt_map[remote.lower()]
    params["start"] = str(max(0, (page - 1) * 10))
    url = f"{LINKEDIN_SEARCH_URL}?{urllib.parse.urlencode(params)}"

    headers = {
        "User-Agent": LINKEDIN_UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
    }
    try:
        code, raw_html = _http_get_with_retry(url, headers=headers, max_retries=2, timeout=12.0)
        if code != 200 or not raw_html:
            return {"results": [], "error": f"LinkedIn returned HTTP {code}"}
        cards = parse_linkedin_cards(raw_html)[:limit]
        return {"results": cards, "error": None}
    except Exception as exc:
        return {"results": [], "error": str(exc)}


def _extract_div_content(raw_html: str, class_name: str) -> str | None:
    pattern = re.compile(rf'<div[^>]*class="[^"]*{re.escape(class_name)}[^"]*"[^>]*>', re.I)
    m = pattern.search(raw_html)
    if not m:
        return None
    start = m.end()
    i = start
    depth = 1
    while depth > 0 and i < len(raw_html):
        next_open = raw_html.find("<div", i)
        next_close = raw_html.find("</div>", i)
        if next_close == -1:
            return None
        if next_open != -1 and next_open < next_close:
            depth += 1
            i = next_open + 4
        else:
            depth -= 1
            i = next_close + 6
    return raw_html[start : i - 6]


def fetch_linkedin_detail(job_id_or_url: str) -> dict:
    job_id = normalize_linkedin_id(job_id_or_url)
    if not job_id:
        return {"ok": False, "error": f"Could not parse LinkedIn job ID from {job_id_or_url!r}"}
    url = f"{LINKEDIN_DETAIL_URL}/{job_id}"
    headers = {
        "User-Agent": LINKEDIN_UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "X-Requested-With": "XMLHttpRequest",
    }
    code, raw_html = _http_get_with_retry(url, headers=headers, max_retries=2, timeout=12.0)
    if code == 404 or not raw_html:
        return {"ok": False, "error": f"LinkedIn job posting not found (HTTP {code})"}

    title_m = re.search(r'class="(?:top-card-layout__title|topcard__title)[^"]*"[^>]*>([\s\S]*?)</h[12]>', raw_html, re.I)
    title = _clean_fragment(title_m.group(1)) if title_m else "(untitled)"

    org_m = re.search(r'class="topcard__org-name-link[^"]*"[^>]*href="([^"]+)"[^>]*>([\s\S]*?)</a>', raw_html, re.I)
    company = _clean_fragment(org_m.group(2)) if org_m else ""
    if not company:
        org_alt = re.search(r'class="topcard__flavor"[^>]*>([\s\S]*?)</span>', raw_html, re.I)
        company = _clean_fragment(org_alt.group(1)) if org_alt else "Target Company"

    loc_m = re.search(r'class="topcard__flavor topcard__flavor--bullet"[^>]*>([\s\S]*?)</span>', raw_html, re.I)
    location = _clean_fragment(loc_m.group(1)) if loc_m else ""

    desc_start = re.search(r'class="(?:show-more-less-html__markup|description__text)', raw_html, re.I)
    topcard = raw_html[: desc_start.start()] if desc_start else raw_html
    is_active = not bool(re.search(r"closed-job__flavor|no longer accepting applications", topcard, re.I))

    desc_html = _extract_div_content(raw_html, "show-more-less-html__markup") or _extract_div_content(raw_html, "description__text") or ""
    cleaned_desc, findings = posting_guard.guard(desc_html or raw_html, "html")

    criteria = {}
    for cm in re.finditer(
        r'class="description__job-criteria-subheader"[^>]*>([\s\S]*?)</h3>[\s\S]*?class="description__job-criteria-text[^"]*"[^>]*>([\s\S]*?)</span>',
        raw_html,
        re.I,
    ):
        criteria[_clean_fragment(cm.group(1)).lower()] = _clean_fragment(cm.group(2))

    return {
        "ok": True,
        "id": job_id,
        "title": title,
        "company": company,
        "location": location,
        "url": f"https://www.linkedin.com/jobs/view/{job_id}",
        "portal": "linkedin-search",
        "is_active": is_active,
        "seniority": criteria.get("seniority level"),
        "employment_type": _normalize_employment_type(criteria.get("employment type")),
        "description": cleaned_desc,
        "guard_verdict": findings.get("verdict", "posting_guard: clean"),
        "guard_findings": findings,
    }


def _normalize_employment_type(raw: str | None) -> str | None:
    if not raw:
        return None
    s = raw.lower()
    if "full" in s or "permanent" in s:
        return "full_time"
    if "part" in s:
        return "part_time"
    if "contract" in s or "fixed" in s:
        return "contract"
    if "freelance" in s:
        return "freelance"
    return None


def search_freehire(query: str, jobage_days: int = 14, limit: int = 25, countries: list[str] | None = None, regions: list[str] | None = None, work_mode: str | None = None) -> dict:
    params: list[tuple[str, str]] = [
        ("limit", str(limit)),
        ("offset", "0"),
        ("description_format", "text"),
    ]
    if query:
        params.append(("q", query))
    if jobage_days and 0 < jobage_days < 9999:
        params.append(("posted_within_days", str(jobage_days)))
    if work_mode:
        params.append(("work_mode", work_mode))
    for c in countries or []:
        if c:
            params.append(("countries", c))
    for r in regions or []:
        if r:
            params.append(("regions", r))

    url = f"{FREEHIRE_BASE_URL}/api/v1/agent/jobs/search?{urllib.parse.urlencode(params)}"
    headers = {"User-Agent": FREEHIRE_UA, "Accept": "application/json"}
    try:
        code, body = _http_get_with_retry(url, headers=headers, max_retries=2, timeout=12.0)
        if code != 200 or not body:
            return {"results": [], "error": f"Freehire returned HTTP {code}"}
        payload = json.loads(body)
        rows = payload.get("data") or []
        out = []
        for j in rows:
            raw_desc = j.get("description") or ""
            cleaned_desc, findings = posting_guard.guard(raw_desc, "auto") if raw_desc else ("", {"verdict": "posting_guard: clean"})
            posted_raw = j.get("posted_at") or ""
            posted_date = posted_raw[:10] if re.match(r"^\d{4}-\d{2}-\d{2}", posted_raw) else None
            enr = j.get("enrichment") or {}
            out.append({
                "id": j.get("public_slug") or j.get("external_id") or "",
                "title": html_lib.unescape(j.get("title") or "").strip() or "(untitled)",
                "company": normalize_company_display(html_lib.unescape(j.get("company") or "").strip()) or "Unknown Company",
                "location": html_lib.unescape(j.get("location") or "").strip(),
                "date": posted_date,
                "url": j.get("url") or "",
                "portal": "freehire-search",
                "source": "cli",
                "work_mode": j.get("work_mode"),
                "employment_type": _normalize_employment_type(enr.get("employment_type")),
                "seniority": enr.get("seniority"),
                "skills": j.get("skills") or [],
                "description": cleaned_desc,
                "guard_verdict": findings.get("verdict", "posting_guard: clean"),
            })
        return {"results": out, "error": None}
    except Exception as exc:
        return {"results": [], "error": str(exc)}


def fetch_freehire_detail(slug_or_url: str) -> dict:
    slug = normalize_freehire_slug(slug_or_url)
    if not slug:
        return {"ok": False, "error": f"Could not parse Freehire slug from {slug_or_url!r}"}
    url = f"{FREEHIRE_BASE_URL}/api/v1/jobs/{urllib.parse.quote(slug)}"
    code, body = _http_get_with_retry(url, headers={"User-Agent": FREEHIRE_UA, "Accept": "application/json"}, max_retries=2, timeout=12.0)
    if code != 200 or not body:
        return {"ok": False, "error": f"Freehire job not found (HTTP {code})"}
    env = json.loads(body)
    j = env.get("data") or {}
    cleaned_desc, findings = posting_guard.guard(j.get("description") or "", "html")
    enr = j.get("enrichment") or {}
    return {
        "ok": True,
        "id": j.get("public_slug") or slug,
        "title": html_lib.unescape(j.get("title") or "").strip() or "(untitled)",
        "company": html_lib.unescape(j.get("company") or "").strip() or "Target Company",
        "location": html_lib.unescape(j.get("location") or "").strip(),
        "url": j.get("url") or f"{FREEHIRE_BASE_URL}/jobs/{slug}",
        "portal": "freehire-search",
        "is_active": True,
        "seniority": enr.get("seniority"),
        "employment_type": _normalize_employment_type(enr.get("employment_type")),
        "description": cleaned_desc,
        "guard_verdict": findings.get("verdict", "posting_guard: clean"),
        "guard_findings": findings,
    }


def extract_greenhouse_target(url: str) -> tuple[str, str] | None:
    """Return (board_token, job_id) for Greenhouse hosted or gh_jid-embedded URLs."""
    try:
        parsed = urllib.parse.urlsplit((url or "").strip())
    except Exception:
        return None
    host = (parsed.hostname or "").lower()
    qs = urllib.parse.parse_qs(parsed.query)
    if "greenhouse.io" in host:
        m = re.match(r"^/([^/]+)/jobs/(\d+)", parsed.path or "")
        if m:
            return m.group(1), m.group(2)
        tok = (qs.get("token") or qs.get("gh_jid") or [""])[0]
        board = (qs.get("for") or [""])[0]
        if tok and board and tok.isdigit():
            return board, tok
    gh_jid = (qs.get("gh_jid") or [""])[0]
    if gh_jid.isdigit():
        board = (qs.get("for") or [""])[0]
        if not board:
            for dom, b in KNOWN_GREENHOUSE_DOMAIN_BOARDS.items():
                if host == dom or host.endswith("." + dom):
                    board = b
                    break
        if board:
            return board, gh_jid
    return None


def fetch_greenhouse_detail(board: str, job_id: str, original_url: str) -> dict:
    api_url = f"https://boards-api.greenhouse.io/v1/boards/{urllib.parse.quote(board)}/jobs/{urllib.parse.quote(job_id)}"
    code, body = _http_get_with_retry(api_url, headers={"User-Agent": FREEHIRE_UA, "Accept": "application/json"}, max_retries=2, timeout=12.0)
    if code == 404:
        return {"ok": False, "is_active": False, "error": f"Posting appears closed or expired at Greenhouse (HTTP 404) for {board}/{job_id}."}
    if code != 200 or not body:
        return {"ok": False, "error": f"Greenhouse API returned HTTP {code}"}
    j = json.loads(body)
    raw_content = html_lib.unescape(j.get("content") or "")
    cleaned_desc, findings = posting_guard.guard(raw_content, "html")
    loc = ((j.get("location") or {}).get("name") or "").strip()
    title = html_lib.unescape(j.get("title") or "").strip() or "(untitled)"
    comp = normalize_company_display(j.get("company_name") or board)
    return {
        "ok": True,
        "id": str(j.get("id") or job_id),
        "title": title,
        "company": comp,
        "location": loc,
        "url": j.get("absolute_url") or original_url,
        "portal": "greenhouse-api",
        "is_active": True,
        "description": cleaned_desc[:18000],
        "guard_verdict": findings.get("verdict", "posting_guard: clean"),
        "guard_findings": findings,
    }


def extract_lever_target(url: str) -> tuple[str, str] | None:
    try:
        parsed = urllib.parse.urlsplit((url or "").strip())
    except Exception:
        return None
    host = (parsed.hostname or "").lower()
    if host == "jobs.lever.co" or host.endswith(".lever.co"):
        parts = [p for p in parsed.path.split("/") if p]
        if len(parts) >= 2 and re.fullmatch(r"[0-9a-f-]{8,}", parts[1], re.I):
            return parts[0], parts[1]
    return None


def fetch_lever_detail(board: str, job_id: str, original_url: str) -> dict:
    api_url = f"https://api.lever.co/v0/postings/{urllib.parse.quote(board)}/{urllib.parse.quote(job_id)}"
    code, body = _http_get_with_retry(api_url, headers={"User-Agent": FREEHIRE_UA, "Accept": "application/json"}, max_retries=2, timeout=12.0)
    if code == 404:
        return {"ok": False, "is_active": False, "error": "Posting appears closed or expired at Lever (HTTP 404)."}
    if code != 200 or not body:
        return {"ok": False, "error": f"Lever API returned HTTP {code}"}
    j = json.loads(body)
    blocks = [j.get("description") or ""]
    for lst in j.get("lists") or []:
        blocks.append(f"<h3>{html_lib.escape(lst.get('text') or '')}</h3>{lst.get('content') or ''}")
    blocks.append(j.get("additional") or "")
    cleaned_desc, findings = posting_guard.guard("\n".join(blocks), "html")
    cats = j.get("categories") or {}
    return {
        "ok": True,
        "id": str(j.get("id") or job_id),
        "title": html_lib.unescape(j.get("text") or "").strip() or "(untitled)",
        "company": normalize_company_display(board),
        "location": (cats.get("location") or "").strip(),
        "url": j.get("hostedUrl") or original_url,
        "portal": "lever-api",
        "is_active": True,
        "description": cleaned_desc[:18000],
        "guard_verdict": findings.get("verdict", "posting_guard: clean"),
        "guard_findings": findings,
    }


def extract_ashby_target(url: str) -> tuple[str, str] | None:
    try:
        parsed = urllib.parse.urlsplit((url or "").strip())
    except Exception:
        return None
    host = (parsed.hostname or "").lower()
    if host == "jobs.ashbyhq.com" or host.endswith(".ashbyhq.com"):
        parts = [p for p in parsed.path.split("/") if p]
        if len(parts) >= 2:
            return parts[0], parts[1]
    return None


def fetch_ashby_detail(board: str, job_id: str, original_url: str) -> dict:
    api_url = f"https://api.ashbyhq.com/posting-api/job-board/{urllib.parse.quote(board)}"
    code, body = _http_get_with_retry(api_url, headers={"User-Agent": FREEHIRE_UA, "Accept": "application/json"}, max_retries=2, timeout=12.0)
    if code != 200 or not body:
        return {"ok": False, "error": f"Ashby API returned HTTP {code}"}
    payload = json.loads(body)
    match = None
    for j in payload.get("jobs") or []:
        if str(j.get("id") or "") == job_id or job_id in (j.get("jobUrl") or ""):
            match = j
            break
    if not match:
        return {"ok": False, "is_active": False, "error": "Posting appears closed or not listed on Ashby board."}
    raw_desc = match.get("descriptionHtml") or match.get("descriptionPlain") or ""
    fmt = "html" if match.get("descriptionHtml") else "text"
    cleaned_desc, findings = posting_guard.guard(raw_desc, fmt)
    return {
        "ok": True,
        "id": str(match.get("id") or job_id),
        "title": html_lib.unescape(match.get("title") or "").strip() or "(untitled)",
        "company": normalize_company_display(board),
        "location": (match.get("location") or "").strip(),
        "url": match.get("jobUrl") or original_url,
        "portal": "ashby-api",
        "is_active": True,
        "description": cleaned_desc[:18000],
        "guard_verdict": findings.get("verdict", "posting_guard: clean"),
        "guard_findings": findings,
    }


def fetch_and_guard_url(url: str) -> dict:
    """Fetch an arbitrary posting URL cleanly via portal/ATS JSON handler or fetch_page + posting_guard."""
    target = (url or "").strip()
    if not target:
        return {"ok": False, "error": "URL is required"}
    ok, reason = is_safe_public_url(target)
    if not ok:
        return {"ok": False, "error": reason}

    if normalize_linkedin_id(target):
        res = fetch_linkedin_detail(target)
        if res.get("ok"):
            return res

    if normalize_freehire_slug(target):
        res = fetch_freehire_detail(target)
        if res.get("ok"):
            return res

    gh = extract_greenhouse_target(target)
    if gh:
        res = fetch_greenhouse_detail(gh[0], gh[1], target)
        if res.get("ok") or res.get("is_active") is False:
            return res

    lv = extract_lever_target(target)
    if lv:
        res = fetch_lever_detail(lv[0], lv[1], target)
        if res.get("ok") or res.get("is_active") is False:
            return res

    ab = extract_ashby_target(target)
    if ab:
        res = fetch_ashby_detail(ab[0], ab[1], target)
        if res.get("ok") or res.get("is_active") is False:
            return res

    # Generic career page / ATS fetch respecting robots_check + fetch_page + posting_guard
    try:
        rc, robots_msg = robots_gate(target)
        if rc != 0:
            return {"ok": False, "error": f"Refused by site robots.txt policy ({robots_msg}). Please paste the JD text manually."}

        fd, tmp_name = tempfile.mkstemp(prefix="job_fetch_", suffix=".html")
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            current = target
            code = 0
            for _ in range(5):
                code, loc = fetch_with_urllib(current, tmp_path, timeout=20.0)
                if code in (301, 302, 303, 307, 308) and loc:
                    next_url = urllib.parse.urljoin(current, loc)
                    ok_next, reason_next = is_safe_public_url(next_url)
                    if not ok_next:
                        return {"ok": False, "error": f"Redirect blocked: {reason_next}"}
                    rc_next, msg_next = robots_gate(next_url)
                    if rc_next != 0:
                        return {"ok": False, "error": f"Redirect refused by robots.txt ({msg_next})"}
                    current = next_url
                    continue
                break

            if code >= 400 or not tmp_path.exists() or tmp_path.stat().st_size == 0:
                return {"ok": False, "error": f"Target returned HTTP {code}. Paste the JD text directly instead."}

            raw_bytes = tmp_path.read_bytes()
            raw_html = decode_page(raw_bytes)
        finally:
            tmp_path.unlink(missing_ok=True)

        # Extract <title> or <h1> hint
        title_m = re.search(r"<h1[^>]*>([\s\S]*?)</h1>", raw_html, re.I) or re.search(r"<title[^>]*>([\s\S]*?)</title>", raw_html, re.I)
        guessed_title = _clean_fragment(title_m.group(1))[:120] if title_m else ""
        cleaned_text, findings = posting_guard.guard(raw_html, "html")

        # Try extracting company from domain if unknown
        host = (urllib.parse.urlsplit(target).hostname or "").lower()
        guessed_company = ""
        if "greenhouse.io" in host or "lever.co" in host or "ashbyhq.com" in host:
            parts = [p for p in urllib.parse.urlsplit(target).path.split("/") if p]
            if parts:
                guessed_company = parts[0].replace("-", " ").title()
        else:
            dom = host.replace("careers.", "").replace("jobs.", "").replace("www.", "").split(".")[0]
            if dom:
                guessed_company = dom.title()

        return {
            "ok": True,
            "id": "",
            "title": guessed_title,
            "company": guessed_company,
            "location": "",
            "url": target,
            "portal": "web-fetch",
            "is_active": True,
            "description": cleaned_text[:18000],
            "guard_verdict": findings.get("verdict", "posting_guard: clean"),
            "guard_findings": findings,
        }
    except Exception as exc:
        return {"ok": False, "error": f"Fetch failed: {exc}"}


def build_referral_links(company: str, title: str) -> dict:
    comp = _company_search_term(company or "")
    cleaned_title = re.sub(r"[,/|–—()]+", " ", title or "Product Manager")
    role_kw = " ".join(cleaned_title.split()[:4])
    rec_q = urllib.parse.quote(f"{comp} recruiter".strip())
    peer_q = urllib.parse.quote(f"{comp} {role_kw}".strip())
    return {
        "recruiter_search_url": f"https://www.linkedin.com/search/results/people/?keywords={rec_q}&origin=GLOBAL_SEARCH_HEADER",
        "peer_search_url": f"https://www.linkedin.com/search/results/people/?keywords={peer_q}&origin=GLOBAL_SEARCH_HEADER",
    }


def load_seen_state() -> tuple[dict, dict]:
    if not SEEN_STATE_PATH.exists():
        SEEN_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        doc = {"seen": {}}
        save_state(SEEN_STATE_PATH, doc)
        return doc, doc["seen"]
    doc = json.loads(SEEN_STATE_PATH.read_text(encoding="utf-8"))
    seen = doc.get("seen") if isinstance(doc, dict) and "seen" in doc else doc
    if not isinstance(seen, dict):
        doc = {"seen": {}}
        seen = doc["seen"]
    return doc, seen


def store_scraped_jobs(hits: list[dict]) -> dict:
    """Deduplicate against seen_jobs.json and tracker.csv using tools/scrape_state & job_key."""
    doc, seen = load_seen_state()
    excluded = tracker_pairs(TRACKER_PATH)
    selected, counts = scrape_select(hits, seen, excluded, strict=False)
    today_iso = datetime.now().strftime("%Y-%m-%d")
    added_keys = []

    for job in selected:
        key = make_key(job["company"], job["title"], job["url"])
        if key in seen:
            continue
        entry = {
            "title": job["title"],
            "company": job["company"],
            "url": job["url"],
            "location": job.get("location") or "",
            "first_seen": today_iso,
            "posted_date": job.get("date") or job.get("posted_date"),
            "deadline": job.get("deadline"),
            "employment_type": job.get("employment_type"),
            "fit": job.get("fit") or "medium",
            "status": job.get("status") or "new",
            "portal": job.get("portal") or "websearch",
            "source": job.get("source") or "cli",
        }
        if job.get("guard_verdict"):
            gv = job["guard_verdict"].strip()
            if gv.lower().startswith("posting_guard:"):
                gv = gv[len("posting_guard:"):].strip()
            entry["guard"] = "clean" if gv.lower().startswith("clean") else gv[:200]
        if job.get("description"):
            entry["description_excerpt"] = job["description"][:2500]
        seen[key] = entry
        added_keys.append(key)

    if added_keys:
        SEEN_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        save_state(SEEN_STATE_PATH, doc)

    return {
        "added_count": len(added_keys),
        "added_keys": added_keys,
        "counts": counts,
    }


def record_rank_scores(results: list[dict]) -> dict:
    """Persist ranked scores into seen_jobs.json using rank_state weights/bands."""
    doc, seen = load_seen_state()
    today_iso = datetime.now().strftime("%Y-%m-%d")
    updated = []
    for res in results:
        key = res.get("key")
        entry = seen.get(key)
        if not entry:
            continue
        if res.get("status") == "expired":
            entry["status"] = "expired"
            updated.append({"key": key, "status": "expired"})
            continue
        scores = res.get("scores") or {}
        total = overall_score(scores)
        verdict_label = band(total)
        entry["status"] = "ranked"
        entry["rank_score"] = total
        entry["rank_verdict"] = verdict_label
        entry["rank_date"] = today_iso
        entry["location_verdict"] = res.get("location_verdict") or "PASS"
        entry["language_gate"] = res.get("language_gate") or "PASS"
        if res.get("language_note") and entry["language_gate"] != "PASS":
            entry["language_note"] = res["language_note"]
        else:
            entry.pop("language_note", None)
        if isinstance(res.get("strengths"), list):
            entry["strengths"] = [str(s) for s in res["strengths"]][:3]
        if isinstance(res.get("gaps"), list):
            entry["gaps"] = [str(g) for g in res["gaps"]][:3]
        if res.get("guard"):
            gv = str(res["guard"]).strip()
            if gv.lower().startswith("posting_guard:"):
                gv = gv[len("posting_guard:"):].strip()
            entry["guard"] = "clean" if gv.lower().startswith("clean") else gv[:200]
        updated.append({"key": key, "score": total, "verdict": verdict_label})

    if updated:
        save_state(SEEN_STATE_PATH, doc)
    return {"updated_count": len(updated), "updated": updated}
