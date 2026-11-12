import re
import urllib.parse


# Seniority levels ordered cleanly for multiplier / tier lookup
# Levels: "apm", "pm", "senior", "staff", "director"
ROLE_LEVEL_LABELS = {
    "apm": "APM / Early-Career PM",
    "pm": "PM / PM II",
    "senior": "Senior / Lead PM",
    "staff": "Staff / Principal PM",
    "director": "Assoc. Director / Director / VP",
}

# Verified 2025-2026 market ranges (total annual CTC in INR Lakhs for India)
# Anchored on AmbitionBox, Glassdoor India, Levels.fyi & peer disclosures.
# tuple format: (low_lakhs, high_lakhs, base_pct, note)
INDIA_COMPANY_BENCHMARKS = {
    "stripe": {
        "apm": (35, 55, 65, "High RSU weighting"),
        "pm": (65, 95, 60, "Base ~₹45-58L + RSUs"),
        "senior": (90, 130, 58, "Base ~₹55-72L + RSUs"),
        "staff": (110, 165, 55, "Base ~₹68-90L + heavy RSUs"),
        "director": (160, 230, 52, "Executive base + RSUs"),
        "tier": "Global Fintech Tier-1",
    },
    "google": {
        "apm": (28, 42, 65, "Base + bonus + GSU"),
        "pm": (50, 78, 60, "Base + 15% bonus + GSUs"),
        "senior": (78, 125, 55, "L6 PM: Base + GSUs"),
        "staff": (120, 185, 52, "L7 PM: Heavy GSU refreshers"),
        "director": (175, 280, 48, "L8 Product Director TC"),
        "tier": "BigTech (FAANG)",
    },
    "microsoft": {
        "apm": (22, 34, 70, "Base + bonus + RSUs"),
        "pm": (38, 62, 65, "L61-62 PM Total Comp"),
        "senior": (58, 92, 62, "L63-64 Senior PM TC"),
        "staff": (88, 140, 58, "L65 Principal PM TC"),
        "director": (130, 200, 55, "L66-67 Partner/Director"),
        "tier": "BigTech",
    },
    "amazon": {
        "apm": (24, 36, 72, "Base + RSU vest"),
        "pm": (42, 68, 65, "L5-L6 PM-T Total Comp"),
        "senior": (65, 105, 60, "L6 Sr PM-T Total Comp"),
        "staff": (105, 165, 55, "L7 Principal PM-T"),
        "director": (150, 240, 50, "L8 Director Product"),
        "tier": "BigTech (FAANG)",
    },
    "meta": {
        "apm": (30, 45, 65, "Base + RSUs"),
        "pm": (55, 85, 58, "IC4/5 Total Comp"),
        "senior": (85, 135, 55, "IC5/6 Total Comp"),
        "staff": (130, 195, 50, "IC6/7 Staff PM"),
        "director": (180, 280, 48, "Director Product"),
        "tier": "BigTech (FAANG)",
    },
    "paypal": {
        "apm": (20, 30, 78, "Structured base + RSU"),
        "pm": (35, 55, 75, "T24-T25 Base + RSU"),
        "senior": (62, 95, 70, "T26 Senior PM: ~₹45-62L base + RSU"),
        "staff": (95, 145, 65, "T27 Staff/Group PM (~₹70L base + RSU)"),
        "director": (135, 195, 60, "T28 Director Product"),
        "tier": "Global Payments Leader",
    },
    "visa": {
        "apm": (18, 28, 80, "Structured band + bonus"),
        "pm": (34, 52, 78, "I5-I6 Base + bonus + RSU"),
        "senior": (55, 80, 76, "I6 Sr PM: ₹45-60L base + bonus + stock"),
        "staff": (80, 120, 72, "I7 Principal/Director band"),
        "director": (110, 165, 68, "I8 Sr Director Product"),
        "tier": "Global Card Network",
    },
    "mastercard": {
        "apm": (18, 28, 80, "Structured band + bonus"),
        "pm": (34, 50, 78, "Base + bonus + LTI"),
        "senior": (52, 78, 76, "Director/Sr PM India band"),
        "staff": (78, 115, 72, "VP / Principal PM band"),
        "director": (110, 160, 68, "SVP / Group Lead band"),
        "tier": "Global Card Network",
    },
    "american express": {
        "apm": (18, 28, 80, "Base + annual bonus"),
        "pm": (32, 48, 78, "Band 35-40 Product"),
        "senior": (48, 75, 76, "Band 40/45 Sr PM / Dir"),
        "staff": (72, 108, 74, "Director Product"),
        "director": (105, 155, 70, "VP Product"),
        "tier": "Global Network & Issuer",
    },
    "adyen": {
        "apm": (22, 34, 78, "Base + equity"),
        "pm": (42, 66, 75, "Base + Adyen equity"),
        "senior": (70, 110, 72, "EU fintech parity in Bengaluru"),
        "staff": (105, 150, 68, "Staff / Principal Product"),
        "director": (140, 200, 65, "Product Lead / Director"),
        "tier": "Global Acquiring Tier-1",
    },
    "phonepe": {
        "apm": (16, 24, 78, "Base + ESOPs"),
        "pm": (36, 56, 72, "Base ~₹30-42L + ESOPs"),
        "senior": (70, 95, 68, "Base ~₹45-60L + ₹20-35L ESOPs"),
        "staff": (100, 145, 62, "Principal/Staff + heavy ESOPs"),
        "director": (135, 195, 58, "Director/VP Product"),
        "tier": "India Fintech Tier-1",
    },
    "razorpay": {
        "apm": (15, 24, 80, "Base + ESOPs"),
        "pm": (30, 48, 76, "PM II: ₹26-38L base + ESOPs"),
        "senior": (48, 76, 74, "Sr PM: ₹40-60L base + ESOPs"),
        "staff": (78, 115, 68, "Staff PM: base + ESOPs"),
        "director": (105, 155, 64, "Director / Group PM"),
        "tier": "India Fintech Tier-1",
    },
    "ebay": {
        "apm": (16, 25, 80, "Base + RSUs"),
        "pm": (30, 46, 76, "Base + bonus + RSUs"),
        "senior": (45, 75, 72, "Base ~₹35-48L + bonus + RSUs"),
        "staff": (75, 115, 68, "Principal / Group PM"),
        "director": (105, 150, 65, "Director Product"),
        "tier": "Global Commerce Platform",
    },
    "payu": {
        "apm": (12, 20, 84, "Base + variable"),
        "pm": (22, 36, 82, "PM II: ₹20-30L base + variable"),
        "senior": (32, 52, 80, "Sr PM: ₹28-44L base + variable"),
        "staff": (52, 78, 76, "Principal / Lead PM"),
        "director": (72, 110, 72, "Director Product"),
        "tier": "India Payments Processor",
    },
    "global payments": {
        "apm": (12, 18, 85, "Structured MNC base + bonus"),
        "pm": (22, 34, 84, "Base + annual bonus"),
        "senior": (32, 52, 82, "Proposition/Sr PM: predictable cash comp"),
        "staff": (48, 72, 80, "Principal / Lead Product"),
        "director": (68, 100, 76, "Director Product"),
        "tier": "Global MNC Processor",
    },
    "juspay": {
        "apm": (14, 22, 80, "Base + ESOPs"),
        "pm": (25, 42, 78, "Core payment infra PM"),
        "senior": (38, 65, 74, "Senior / Platform PM + ESOPs"),
        "staff": (62, 95, 70, "Staff / Principal PM"),
        "director": (85, 130, 66, "Director / Head of Product"),
        "tier": "Payments Infra Unicorn",
    },
    "cashfree": {
        "apm": (14, 22, 80, "Base + ESOPs"),
        "pm": (24, 40, 78, "API banking PM"),
        "senior": (38, 65, 74, "Senior PM: ₹32-52L base + ESOPs"),
        "staff": (60, 92, 70, "Principal / Group PM"),
        "director": (85, 125, 66, "Director / VP Product"),
        "tier": "India Payments Unicorn",
    },
    "pine labs": {
        "apm": (12, 19, 82, "Base + ESOPs"),
        "pm": (22, 35, 80, "POS / merchant issuing PM"),
        "senior": (34, 56, 76, "Sr PM: ₹28-46L base + ESOPs"),
        "staff": (54, 82, 72, "Principal / Group PM"),
        "director": (75, 115, 68, "Director / AVP Product"),
        "tier": "POS & Merchant Payments",
    },
    "airwallex": {
        "apm": (20, 30, 75, "Cross-border fintech TC"),
        "pm": (36, 56, 72, "Base + equity"),
        "senior": (55, 88, 68, "Global fintech premium + equity"),
        "staff": (85, 130, 64, "Staff / Principal PM"),
        "director": (120, 175, 60, "Director Product"),
        "tier": "Global Cross-Border Fintech",
    },
    "xendit": {
        "apm": (16, 25, 78, "Regional SEA fintech band"),
        "pm": (28, 46, 75, "Base + equity"),
        "senior": (45, 75, 72, "Regional payments platform PM"),
        "staff": (70, 108, 68, "Staff / Principal PM"),
        "director": (95, 140, 65, "Director / Group PM"),
        "tier": "SEA Payments Unicorn",
    },
    "gokwik": {
        "apm": (14, 21, 82, "Base + ESOPs"),
        "pm": (24, 38, 80, "Checkout & conversion PM"),
        "senior": (36, 58, 76, "Growth fintech: ₹30-46L base + ESOPs"),
        "staff": (55, 85, 72, "Lead / Principal PM"),
        "director": (75, 115, 68, "Director Product"),
        "tier": "E-commerce Checkout Fintech",
    },
    "circle": {
        "apm": (20, 30, 75, "Web3/stablecoin fintech"),
        "pm": (38, 58, 72, "Base + equity"),
        "senior": (58, 92, 68, "Global USDC infra pay scale"),
        "staff": (85, 135, 64, "Principal PM Core Network"),
        "director": (120, 180, 60, "Director Product"),
        "tier": "Global Stablecoin Infra",
    },
    "bitgo": {
        "apm": (18, 28, 76, "Digital assets custody"),
        "pm": (34, 54, 72, "Base + equity"),
        "senior": (52, 85, 68, "Custody / tokenization fintech scale"),
        "staff": (80, 125, 64, "Principal / Staff PM"),
        "director": (110, 165, 60, "Director Product"),
        "tier": "Digital Asset Custody",
    },
    "nium": {
        "apm": (15, 23, 80, "Cross-border rails"),
        "pm": (26, 42, 78, "Base + ESOPs"),
        "senior": (44, 70, 74, "Global B2B payout fintech"),
        "staff": (68, 102, 70, "Staff PM Digital Assets"),
        "director": (92, 138, 66, "Director / VP Product"),
        "tier": "Cross-Border Payments",
    },
    "commonwealth bank": {
        "apm": (15, 23, 84, "Global bank GCC band"),
        "pm": (26, 42, 82, "Structured GCC base + bonus"),
        "senior": (42, 66, 80, "Stream / Sr Product Lead GCC"),
        "staff": (62, 92, 78, "Principal / Lead GCC"),
        "director": (85, 125, 75, "Executive Manager GCC"),
        "tier": "Global Bank GCC (India)",
    },
    "societe generale": {
        "apm": (14, 21, 85, "Global bank GCC band"),
        "pm": (22, 36, 84, "Product Owner / PM GCC"),
        "senior": (34, 54, 82, "Lead PO / Senior PM GCC"),
        "staff": (50, 75, 80, "Principal / Chapter Lead"),
        "director": (70, 105, 78, "Director / VP GCC"),
        "tier": "Global Bank GCC (India)",
    },
    "arcesium": {
        "apm": (18, 26, 80, "D. E. Shaw fintech spinoff"),
        "pm": (32, 50, 78, "High fixed + performance bonus"),
        "senior": (48, 76, 75, "Post-trade fintech TC"),
        "staff": (74, 112, 72, "Principal Product Manager"),
        "director": (105, 155, 68, "Director / SVP Product"),
        "tier": "Institutional FinTech",
    },
    "aspire": {
        "apm": (15, 24, 78, "APAC B2B neobank"),
        "pm": (28, 46, 75, "Cross-border B2B fintech"),
        "senior": (45, 72, 72, "Senior PM FX & Cross-border"),
        "staff": (68, 105, 68, "Staff / Principal PM"),
        "director": (95, 138, 65, "Director / Head of Product"),
        "tier": "APAC B2B Neobank",
    },
    "cleartax": {
        "apm": (14, 21, 80, "B2B SaaS / Fintech"),
        "pm": (24, 40, 78, "Base + ESOPs"),
        "senior": (38, 60, 75, "Base + ESOPs"),
        "staff": (55, 85, 72, "Assoc. Dir / Principal PM"),
        "director": (72, 112, 68, "Assoc. Dir / Director Product"),
        "tier": "India B2B Fintech SaaS",
    },
    "godaddy": {
        "apm": (16, 24, 78, "Global SaaS India GCC"),
        "pm": (28, 45, 75, "Base + RSUs"),
        "senior": (45, 72, 72, "Base + bonus + RSUs"),
        "staff": (70, 108, 68, "Principal PM"),
        "director": (98, 145, 64, "Director Product"),
        "tier": "Global SaaS (India Hub)",
    },
    "redpin": {
        "apm": (12, 19, 82, "Cross-border FX/remittance"),
        "pm": (22, 36, 80, "Base + variable"),
        "senior": (35, 56, 78, "Global FX fintech India hub"),
        "staff": (52, 80, 75, "Principal / Group PM"),
        "director": (72, 108, 72, "Director Product"),
        "tier": "Cross-Border FX Fintech",
    },
    "jpmc": {
        "apm": (16, 24, 82, "Banking tech GCC"),
        "pm": (28, 44, 80, "VP Product band"),
        "senior": (45, 72, 78, "VP / Senior PM Banking Tech"),
        "staff": (68, 105, 75, "ED / Principal PM"),
        "director": (98, 148, 72, "Executive Director / MD"),
        "tier": "Tier-1 Investment Bank",
    },
    "goldman sachs": {
        "apm": (20, 30, 80, "VP/Associate banking tech"),
        "pm": (32, 52, 76, "Base + discretionary bonus"),
        "senior": (52, 84, 74, "VP Product Engineering"),
        "staff": (80, 125, 70, "Senior VP / Principal"),
        "director": (115, 175, 66, "Managing Director Track"),
        "tier": "Tier-1 Investment Bank",
    },
    "payglocal": {
        "apm": (12, 19, 80, "Cross-border payments"),
        "pm": (22, 36, 78, "Base + ESOPs"),
        "senior": (36, 58, 74, "Cross-border acquiring PM"),
        "staff": (55, 85, 70, "Principal / Group PM"),
        "director": (75, 115, 66, "Director / VP Product"),
        "tier": "Cross-Border Fintech",
    },
    "navi": {
        "apm": (14, 22, 80, "Base + ESOPs"),
        "pm": (26, 42, 76, "Consumer fintech PM"),
        "senior": (44, 72, 72, "Senior PM: high ownership"),
        "staff": (70, 108, 68, "Principal / Lead PM"),
        "director": (95, 145, 64, "Director / Group PM"),
        "tier": "India Fintech",
    },
}

DEFAULT_INDIA_BENCHMARK = {
    "apm": (10, 18, 84, "General India Fintech/SaaS APM avg"),
    "pm": (22, 38, 80, "Mid-level PM market median (AmbitionBox/Glassdoor)"),
    "senior": (38, 64, 76, "Senior PM Bengaluru fintech median"),
    "staff": (65, 105, 70, "Staff / Principal PM market median"),
    "director": (85, 140, 66, "Assoc. Dir / Director / VP market median"),
    "tier": "India Tech / Fintech Market Avg",
}

# Non-India Regional Benchmarks (Low, High, Currency, BasePct, Note)
REGIONAL_BENCHMARKS = {
    "US": {
        "apm": (125, 175, "USD", 75, "Base + bonus/RSU"),
        "pm": (165, 235, "USD", 70, "Base + bonus + equity"),
        "senior": (205, 310, "USD", 64, "Top-tier US fintech TC + RSUs"),
        "staff": (275, 430, "USD", 56, "Staff PM TC (base $210-260k + RSU)"),
        "director": (340, 560, "USD", 50, "Director / Group PM TC"),
    },
    "UK": {
        "apm": (60, 90, "GBP", 85, "London fintech base + bonus"),
        "pm": (85, 130, "GBP", 80, "Base + bonus + equity"),
        "senior": (120, 205, "GBP", 74, "London fintech TC (£90-135k base + RSU)"),
        "staff": (160, 265, "GBP", 68, "Staff PM London fintech TC"),
        "director": (200, 340, "GBP", 62, "Director Product London TC"),
    },
    "AE": {
        "apm": (200, 300, "AED", 90, "Tax-free UAE annual package"),
        "pm": (270, 420, "AED", 88, "Tax-free UAE annual package"),
        "senior": (360, 640, "AED", 85, "Tax-free (~$98k-$174k USD equiv)"),
        "staff": (520, 820, "AED", 80, "Tax-free (~$142k-$223k USD equiv)"),
        "director": (650, 1050, "AED", 75, "Tax-free (~$177k-$286k USD equiv)"),
    },
    "SG": {
        "apm": (90, 135, "SGD", 82, "Base + AWS/bonus"),
        "pm": (125, 185, "SGD", 78, "Base + bonus + RSU"),
        "senior": (165, 265, "SGD", 74, "APAC hub TC (~$123k-$198k USD equiv)"),
        "staff": (235, 360, "SGD", 68, "Staff PM SG TC (~$175k-$268k USD equiv)"),
        "director": (290, 460, "SGD", 64, "Director Product SG TC"),
    },
    "IE": {
        "apm": (75, 105, "EUR", 82, "Dublin tech hub TC"),
        "pm": (105, 155, "EUR", 76, "Base + bonus + RSUs"),
        "senior": (140, 225, "EUR", 70, "Dublin Tier-1 tech TC (Stripe/BigTech)"),
        "staff": (190, 300, "EUR", 62, "Staff PM Dublin TC"),
        "director": (240, 380, "EUR", 58, "Director Product Dublin TC"),
    },
    "CA": {
        "apm": (110, 150, "CAD", 80, "Toronto/Remote Canada TC"),
        "pm": (150, 210, "CAD", 75, "Base + bonus + RSUs"),
        "senior": (195, 285, "CAD", 68, "Toronto fintech TC (Stripe/Shopify)"),
        "staff": (260, 380, "CAD", 60, "Staff PM Canada TC"),
        "director": (310, 470, "CAD", 55, "Director Product Canada TC"),
    },
}

USD_CONVERSION_FACTORS = {
    "INR_L": 100000 / 84.5,  # 1 INR Lakh ≈ $1,183 USD
    "USD": 1000.0,
    "GBP": 1270.0,
    "AED": 272.29,
    "SGD": 745.0,
    "EUR": 1080.0,
    "CAD": 720.0,
}


def detect_role_level(title: str) -> str:
    tl = (title or "").lower()
    if any(w in tl for w in ["vice president", "vp ", "vp,", "head of", "director", "group product", "gpm"]):
        return "director"
    if any(w in tl for w in ["senior staff", "principal", "staff product", "staff pm", "stream lead", "product lead", "senior manager"]):
        return "staff" if "principal" in tl or "staff" in tl or "senior staff" in tl else "senior"
    if any(w in tl for w in ["associate director"]):
        return "director"
    if any(w in tl for w in ["associate product", "apm", "new grad", "accelerator", "junior product"]):
        return "apm"
    if any(w in tl for w in ["senior ", "sr.", "sr ", "lead ", "proposition manager", "business manager to ceo"]):
        return "senior"
    if any(w in tl for w in ["ii", "product manager 2", "pm ii", "pm-2", "product owner"]):
        return "pm"
    # Default plain "Product Manager" without Senior/Staff qualifiers
    return "pm"


def detect_region_code(location: str, jd_text: str = "") -> str:
    loc = f"{location or ''}".lower()
    if any(w in loc for w in ["india", "bengaluru", "bangalore", "mumbai", "hyderabad", "pune", "chennai", "gurgaon", "gurugram", "noida", "delhi", "jaipur", "kolkata"]):
        return "IN"
    if any(w in loc for w in ["dubai", "uae", "abu dhabi", "united arab emirates"]):
        return "AE"
    if any(w in loc for w in ["london", "united kingdom", "uk", "manchester"]):
        return "UK"
    if any(w in loc for w in ["singapore"]):
        return "SG"
    if any(w in loc for w in ["dublin", "ireland"]):
        return "IE"
    if any(w in loc for w in ["toronto", "canada", "vancouver"]):
        return "CA"
    if any(w in loc for w in [
        "san francisco", "new york", "new-york", "nyc", "seattle", "sf", "sea",
        "chicago", "atlanta", "us -", "us-", "united states", "remote-us", "remote - us",
    ]) or re.search(r"\bus\b", loc):
        return "US"
    # Fallback using JD text if location was N/A or blank
    jl = (jd_text or "").lower()
    if any(w in jl for w in ["bengaluru", "bangalore", "mumbai", "rbi", "lpa"]):
        return "IN"
    if any(w in jl for w in ["401(k)", "401k", "united states"]):
        return "US"
    return "IN" if not loc.strip() else "US"


def extract_stated_comp(text: str) -> str | None:
    """Extract explicitly advertised compensation strings from JD text when present."""
    if not text:
        return None
    patterns = [
        r"(?:USD|\$|EUR|€|GBP|£|INR|₹|AED|SGD)\s?\d{2,3}(?:,\d{3})+(?:\s?[-–to]+\s?(?:USD|\$|EUR|€|GBP|£|INR|₹|AED|SGD)?\s?\d{2,3}(?:,\d{3})+)?(?:\s*(?:per\s+year|annually|/\s*yr|pa|p\.a\.|base|ote))?",
        r"(?:₹|INR|Rs\.?)\s?\d+(?:\.\d+)?\s*(?:L|LPA|Lakhs?|Cr)(?:\s?[-–to]+\s?\d+(?:\.\d+)?\s*(?:L|LPA|Lakhs?|Cr))?",
        r"\b(?:up\s*to\s*)?\d+(?:\.\d+)?\s*(?:[-–to]+\s*\d+(?:\.\d+)?\s*)?(?:LPA|Lacs?|Lakhs?)\b",
        r"\$\s?\d{2,3}k\s?[-–to]+\s?\$?\d{2,3}k(?:\s*(?:base|tc|total|ote))?",
    ]
    for pat in patterns:
        m = re.search(pat, text, flags=re.I)
        if m:
            return re.sub(r"\s+", " ", m.group(0)).strip()
    return None


def _fmt_inr_lakh(val: float) -> str:
    if val >= 100:
        return f"₹{val / 100:.2f}".rstrip("0").rstrip(".") + " Cr"
    return f"₹{int(round(val))}L"


def _clean_company_for_lookup(company: str) -> str:
    s = re.sub(
        r"\b(?:private\s+limited|pvt\.?\s*ltd\.?|limited|ltd\.?|inc\.?|llc|corp\.?|solutions)\b",
        "",
        company or "",
        flags=re.I,
    )
    return re.sub(r"\s+", " ", s).strip(" ,.-") or (company or "")


def _match_india_company(company: str) -> tuple[str | None, dict]:
    cl = (company or "").strip().lower()
    for name, table in INDIA_COMPANY_BENCHMARKS.items():
        if name in cl:
            return name, table
    return None, DEFAULT_INDIA_BENCHMARK


def build_salary_lookup_links(company: str, title: str, region: str) -> dict:
    clean_comp = _clean_company_for_lookup(company)
    role_short = "Senior Product Manager" if "senior" in (title or "").lower() else "Product Manager"
    q_amb = urllib.parse.quote(f"site:ambitionbox.com/salaries {clean_comp} {role_short}".strip())
    gd_domain = "glassdoor.co.in" if region == "IN" else "glassdoor.com"
    q_gd = urllib.parse.quote(f"site:{gd_domain}/Salary {clean_comp} {role_short}".strip())
    q_lvl = urllib.parse.quote(clean_comp)
    return {
        "glassdoor_url": f"https://www.google.com/search?q={q_gd}",
        "ambitionbox_url": f"https://www.google.com/search?q={q_amb}",
        "levels_url": f"https://www.levels.fyi/companies?search={q_lvl}",
    }


def estimate_market_salary(company: str, title: str, location: str = "", jd_text: str = "") -> dict:
    """
    Returns probable market compensation range & average for a role based on 2025-26
    Glassdoor / AmbitionBox / Levels.fyi benchmarks, plus verbatim JD comp if stated.
    """
    comp_clean = (company or "Target Company").strip()
    level = detect_role_level(title)
    region = detect_region_code(location, jd_text)
    stated = extract_stated_comp(jd_text)

    # AI / Specialized premium adjustment for non-bigtech companies (+10% on AI PM roles)
    tl = f"{title} {jd_text}".lower()
    is_ai_role = bool(re.search(r"\b(ai|genai|llm|agentic|machine learning|ml)\b", (title or "").lower()))

    if region == "IN":
        matched_key, table = _match_india_company(comp_clean)
        low_l, high_l, base_pct, split_note = table[level]
        if is_ai_role and not matched_key:
            low_l = round(low_l * 1.12)
            high_l = round(high_l * 1.15)
            split_note += " (+AI PM premium)"
        avg_l = round((low_l + high_l) / 2.0, 1)
        low_s, high_s, avg_s = _fmt_inr_lakh(low_l), _fmt_inr_lakh(high_l), _fmt_inr_lakh(avg_l)
        range_display = f"{low_s} – {high_s}" + ("" if high_l >= 100 else " LPA")
        avg_display = f"Avg ~{avg_s}" + ("" if avg_l >= 100 else " LPA")
        mid_usd_k = round((avg_l * USD_CONVERSION_FACTORS["INR_L"]) / 1000.0)
        tier_label = table["tier"]
        comp_bucket_high = high_l >= 50
        source_basis = "AmbitionBox / Glassdoor / Levels.fyi (India)"
        if avg_l >= 65:
            expectation_fit = "High / Top Fintech Band"
        elif avg_l >= 42:
            expectation_fit = "Strong / Senior PM Target"
        elif avg_l >= 28:
            expectation_fit = "Mid-Market Band"
        else:
            expectation_fit = "Below Senior PM Target"
    else:
        reg_tbl = REGIONAL_BENCHMARKS.get(region, REGIONAL_BENCHMARKS["US"])
        low_k, high_k, currency, base_pct, split_note = reg_tbl[level]
        # Top fintech premium for Stripe / Google / Meta / Adyen / Coinbase / Airwallex globally
        if any(k in comp_clean.lower() for k in ["stripe", "google", "meta", "anthropic", "coinbase", "adyen", "airwallex"]):
            low_k = round(low_k * 1.12)
            high_k = round(high_k * 1.15)
        avg_k = round((low_k + high_k) / 2.0)
        sym = {"USD": "$", "GBP": "£", "AED": "AED ", "SGD": "S$", "EUR": "€", "CAD": "C$"}.get(currency, "$")
        range_display = f"{sym}{low_k}k – {sym}{high_k}k"
        avg_display = f"Avg ~{sym}{avg_k}k"
        mid_usd_k = round((avg_k * USD_CONVERSION_FACTORS.get(currency, 1000.0)) / 1000.0)
        tier_label = f"{region} Market ({currency})"
        comp_bucket_high = mid_usd_k >= 100
        source_basis = "Levels.fyi / Glassdoor Market Est."
        if mid_usd_k > 200:
            expectation_fit = "Exceeds $200k Target"
        elif mid_usd_k >= 100:
            expectation_fit = "Within $100k–$200k Target"
        else:
            expectation_fit = "Below $100k Global Target"

    links = build_salary_lookup_links(comp_clean, title, region)

    return {
        "stated_in_jd": stated,
        "source_type": "jd_stated" if stated else "market_estimate",
        "source_basis": source_basis,
        "region": region,
        "level_key": level,
        "level_label": ROLE_LEVEL_LABELS.get(level, "Product Manager"),
        "tier_label": tier_label,
        "range_display": range_display,
        "avg_display": avg_display,
        "split_note": split_note,
        "base_pct_est": base_pct,
        "approx_usd_k": mid_usd_k,
        "expectation_fit": expectation_fit,
        "meets_senior_bar": comp_bucket_high,
        **links,
    }
