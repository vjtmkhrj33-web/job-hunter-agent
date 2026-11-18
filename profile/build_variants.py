"""Generates the 3 canonical CV variants (.md + .docx) directly from profile/master_cv.md."""
import re
import sys
from pathlib import Path

PROFILE_DIR = Path(__file__).resolve().parent
BASE_DIR = PROFILE_DIR.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from core.engine import JobAgentEngine

VARIANTS = {
    "master_cv_payments_ai": {
        "headline": "Senior Product Manager | Payments Platforms, Merchant Acquiring and Agentic AI",
        "summary": (
            "Senior Product Manager with almost 14 years in payments technology, nearly 10 in product management, across "
            "merchant acquiring, authorisation, clearing, settlement, tokenisation and agentic AI delivery. At Visa, "
            "contributed to $30M in net revenue by bringing six domestic schemes onto the platform, owned the roadmap for "
            "a self-service settlement platform serving 400+ acquiring banks at up to 65K TPS, and led an agentic "
            "transaction-validation system that prevented $7.5M in high-risk exposure in one week of live data. "
            "Claude-based workflows with human review cut delivery schedules by about 50% on average. Two-time Visa STAR "
            "Award winner, Cognizant Best Employee six years running and Certified Scrum Product Owner."
        ),
    },
    "master_cv_leadership": {
        "headline": "Principal / Group Product Manager | Payments Platforms, Acquiring Scale & Agentic AI",
        "summary": (
            "Payments Product Leader with almost 14 years in financial technology, nearly 10 in product management, "
            "spanning Visa (CyberSource / Visa Platform Connect) and Cognizant (First Data / Fiserv). Led a 12-member "
            "direct team of PMs and analysts (raising post-release accuracy 80% to 97%, lifting monthly certifications "
            "by +50% to +80%) and prioritised RICE roadmaps across four scrum teams (~24 people) with VP sign-off. "
            "Contributed to $30M in net revenue across six domestic schemes, owned a 65K TPS self-service settlement "
            "platform for 400+ acquiring banks, led a 0-to-1 Claude agentic validation system preventing $7.5M exposure "
            "in one week, and built automation turning 300 analyst-hours/day into a 1-hour run (moving 25 of 30 analysts "
            "into Project Lead roles). Two-time Visa STAR Award winner (2024, 2025) and Cognizant Best Employee 2013-2018."
        ),
    },
    "master_cv_generalist": {
        "headline": "Senior Technical Product Manager | Platform Modernisation, APIs & Applied Agentic AI",
        "summary": (
            "Senior Technical Product Manager with almost 14 years in enterprise financial platforms, nearly 10 in product "
            "management, specialising in 0-to-1 applied AI, microservices decomposition, REST/Kafka platform scale, and "
            "regulated delivery. At Visa, led a Claude-based transaction-validation system built on shadow-mode evaluation "
            "frameworks protecting deterministic controls ($7.5M high-risk exposure prevented in one week; p95 <1.2s; "
            "TTFT <350ms), cut delivery schedules by ~50% on average (~10 FTEs tracked in Jira), owned a self-service "
            "platform serving 400+ enterprise banks at up to 65K TPS, and defined a 5-microservice roadmap decoupling ~85% "
            "of a legacy monolith. Two-time Visa STAR Award winner, Cognizant Best Employee 2013-2018 and CSPO."
        ),
    },
}


def build_all_variants():
    master_md = (PROFILE_DIR / "master_cv.md").read_text(encoding="utf-8")
    engine = JobAgentEngine()

    for slug, cfg in VARIANTS.items():
        variant_md = re.sub(
            r"(?m)^(# AVIJIT MUKHERJEE\r?\n)\*\*.*?\*\*$",
            rf"\1**{cfg['headline']}**",
            master_md,
            count=1,
        )
        variant_md = re.sub(
            r"(?s)(## PROFESSIONAL SUMMARY\r?\n\r?\n).*?(\r?\n\r?\n---)",
            rf"\1{cfg['summary']}\2",
            variant_md,
            count=1,
        )

        md_out = PROFILE_DIR / f"{slug}.md"
        docx_out = PROFILE_DIR / f"{slug}.docx"
        md_out.write_text(variant_md, encoding="utf-8")
        engine.generate_docx(variant_md, docx_out)
        print(f"Built variant: {md_out.name} + {docx_out.name}")


if __name__ == "__main__":
    build_all_variants()
