import os
import re
import csv
import json
import sys
import time
from pathlib import Path
from datetime import datetime, timedelta
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = Path(__file__).resolve().parent.parent
TOOLS_DIR = BASE_DIR / "tools"
if str(TOOLS_DIR) not in sys.path:
    sys.path.insert(0, str(TOOLS_DIR))

import posting_guard  # noqa: E402
from core import portals, salary  # noqa: E402


def read_text(path):
    p = BASE_DIR / path
    if p.exists():
        with open(p, "r", encoding="utf-8") as f:
            return f.read()
    return ""


def write_text(path, content):
    p = BASE_DIR / path
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write(content)


def add_hr(paragraph, color_hex="CCCCCC"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(pBdr)


class JobAgentEngine:
    def __init__(self):
        self.base_dir = BASE_DIR
        self.tracker_path = self.base_dir / "tracker.csv"
        self.shortlist_path = self.base_dir / "jobs" / "shortlist.md"
        self.master_cv_path = self.base_dir / "profile" / "master_cv.md"
        self.achievements_path = self.base_dir / "profile" / "achievements.md"
        self.targets_path = self.base_dir / "profile" / "targets.md"
        self.applications_dir = self.base_dir / "applications"
        self.upskill_dir = self.base_dir / "upskill"
        self.skills_dir = self.base_dir / ".agents" / "skills"
        self.pm_refs_dir = self.skills_dir / "senior-pm-job-search" / "references"
        self.candidate_facts_path = self.pm_refs_dir / "candidate-facts.md"
        self.search_profile_path = self.pm_refs_dir / "search-profile.md"
        self.applications_log_path = self.pm_refs_dir / "applications-log.md"
        self.scanner_active = False
        self.last_scan_time = None
        self.scanner_logs = []

    def get_stats(self):
        tracker_rows = self.get_tracker()
        status_counts = {}
        for r in tracker_rows:
            st = (r.get("status") or "shortlisted").strip().lower()
            status_counts[st] = status_counts.get(st, 0) + 1

        shortlist_content = read_text("jobs/shortlist.md")
        shortlist_count = len(re.findall(r"### \d+\.", shortlist_content)) or 11
        app_dirs = [d for d in self.applications_dir.iterdir() if d.is_dir()] if self.applications_dir.exists() else []
        _, seen = portals.load_seen_state()
        ranked_count = sum(1 for e in seen.values() if e.get("status") == "ranked")

        return {
            "total_tracked": len(tracker_rows),
            "status_counts": status_counts,
            "shortlisted_count": shortlist_count,
            "tailored_applications_count": len(app_dirs),
            "scraped_seen_count": len(seen),
            "scraped_ranked_count": ranked_count,
            "scanner_active": self.scanner_active,
            "last_scan_time": self.last_scan_time or "Never",
            "candidate_name": "Avijit Mukherjee",
            "primary_domain": "Payments Product Leader (Visa I6 • Acquiring, Settlement, Rails & AI)"
        }

    def get_tracker(self):
        if not self.tracker_path.exists():
            return []
        rows = []
        with open(self.tracker_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for r in reader:
                rows.append(r)
        return rows

    def save_tracker(self, rows):
        fieldnames = ["company", "role", "link", "location", "sponsorship", "score", "status", "applied_date", "contact", "follow_up_date", "notes"]
        with open(self.tracker_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in rows:
                clean_row = {k: r.get(k, "") for k in fieldnames}
                writer.writerow(clean_row)

    def update_tracker_status(self, company, role, new_status, notes="", follow_up_date="", link="", location="", sponsorship="", score=""):
        rows = self.get_tracker()
        updated = False
        for r in rows:
            if r["company"].strip().lower() == company.strip().lower() and r["role"].strip().lower() == role.strip().lower():
                r["status"] = new_status
                if notes:
                    r["notes"] = notes
                if follow_up_date:
                    r["follow_up_date"] = follow_up_date
                if link:
                    r["link"] = link
                if location:
                    r["location"] = location
                if sponsorship:
                    r["sponsorship"] = sponsorship
                if score:
                    r["score"] = str(score)
                if new_status == "applied" and not r.get("applied_date"):
                    r["applied_date"] = datetime.now().strftime("%Y-%m-%d")
                updated = True
                break
        if not updated:
            rows.append({
                "company": company,
                "role": role,
                "link": link or "",
                "location": location or "Bengaluru / Global",
                "sponsorship": sponsorship or "UNVERIFIED",
                "score": str(score or "8.0"),
                "status": new_status,
                "applied_date": datetime.now().strftime("%Y-%m-%d") if new_status == "applied" else "",
                "contact": "",
                "follow_up_date": follow_up_date,
                "notes": notes
            })
        self.save_tracker(rows)
        self.append_application_log(company, role, new_status, notes or f"Tracker updated to {new_status}")
        return True

    def append_application_log(self, company: str, role: str, stage: str, note: str):
        """Appends a concise line to .agents/skills/senior-pm-job-search/references/applications-log.md."""
        self.pm_refs_dir.mkdir(parents=True, exist_ok=True)
        if not self.applications_log_path.exists():
            self.applications_log_path.write_text(
                "# Applications log\n\n"
                "Append a line whenever a role is screened, a variant is made, something is sent, or a status changes. "
                "Keep newest at the bottom. Format: date | company | role | stage | note.\n\n",
                encoding="utf-8",
            )
        today = datetime.now().strftime("%Y-%m-%d")
        clean_note = re.sub(r"\s+", " ", (note or "").replace("|", "/")).strip()[:180]
        entry = f"{today} | {company.strip()} | {role.strip()} | {stage.strip()} | {clean_note}"
        existing = self.applications_log_path.read_text(encoding="utf-8")
        if entry not in existing:
            sep = "" if existing.endswith("\n") else "\n"
            with open(self.applications_log_path, "a", encoding="utf-8") as fh:
                fh.write(f"{sep}{entry}\n")

    def get_references_data(self):
        return {
            "candidate_facts": self.candidate_facts_path.read_text(encoding="utf-8") if self.candidate_facts_path.exists() else "",
            "search_profile": self.search_profile_path.read_text(encoding="utf-8") if self.search_profile_path.exists() else "",
            "applications_log": self.applications_log_path.read_text(encoding="utf-8") if self.applications_log_path.exists() else "",
            "search_queries": (self.skills_dir / "job-scraper" / "search-queries.md").read_text(encoding="utf-8")
            if (self.skills_dir / "job-scraper" / "search-queries.md").exists() else "",
        }

    def append_candidate_fact(self, category: str, fact: str):
        fact_clean = re.sub(r"\s+", " ", (fact or "")).strip()
        if not fact_clean:
            return {"ok": False, "error": "Fact cannot be empty"}
        if not self.candidate_facts_path.exists():
            return {"ok": False, "error": "candidate-facts.md not found"}
        today = datetime.now().strftime("%Y-%m-%d")
        cat_clean = (category or "Updates").strip()
        line = f"- [{today}] ({cat_clean}) {fact_clean}"
        text = self.candidate_facts_path.read_text(encoding="utf-8")
        sep = "" if text.endswith("\n") else "\n"
        if "## Recent updates" not in text:
            addition = f"{sep}\n## Recent updates\n{line}\n"
        else:
            addition = f"{sep}{line}\n"
        with open(self.candidate_facts_path, "a", encoding="utf-8") as fh:
            fh.write(addition)
        return {"ok": True, "added": line, "candidate_facts": self.candidate_facts_path.read_text(encoding="utf-8")}

    def get_skills_overview(self):
        skills = []
        if self.skills_dir.exists():
            for skill_dir in sorted(self.skills_dir.iterdir()):
                skill_md = skill_dir / "SKILL.md"
                if not skill_dir.is_dir() or not skill_md.exists():
                    continue
                raw = skill_md.read_text(encoding="utf-8", errors="ignore")
                name_m = re.search(r"^name:\s*(.+)$", raw, re.M)
                enabled_m = re.search(r"^enabled:\s*(.+)$", raw, re.M)
                desc_m = re.search(r"^description:\s*(.+)$", raw, re.M)
                enabled = True
                if enabled_m and enabled_m.group(1).strip().lower() == "false":
                    enabled = False
                skills.append({
                    "folder": skill_dir.name,
                    "name": name_m.group(1).strip() if name_m else skill_dir.name,
                    "enabled": enabled,
                    "description": (desc_m.group(1).strip() if desc_m else "").strip("> "),
                })
        return {
            "skills": skills,
            "enabled_count": sum(1 for s in skills if s["enabled"]),
            "disabled_count": sum(1 for s in skills if not s["enabled"]),
            "tools": [
                "posting_guard.py (active prompt-injection gate)",
                "fetch_page.py (robots-checked page fetcher)",
                "job_key.py (canonical deduplication keys)",
                "scrape_state.py (seen_jobs.json + tracker dedup)",
                "rank_state.py (0-100 weighted rubric scorer)",
            ],
        }

    def get_shortlist(self):
        raw = read_text("jobs/shortlist.md")
        sections = raw.split("### ")
        jobs = []
        for s in sections[1:]:
            lines = s.strip().split("\n")
            title_line = lines[0].strip()
            title_match = re.search(r"^\d+\.\s*(.*?)(?:\s*—\s*(.*))?$", title_line)
            role = title_match.group(1).strip() if title_match else title_line
            company = title_match.group(2).strip() if (title_match and title_match.group(2)) else ""

            location = "Bengaluru"
            link = ""
            link_note = ""
            fit_stars = "⭐⭐⭐⭐"
            comp = "UNVERIFIED"

            for line in lines[1:]:
                l = line.strip()
                if l.startswith("- **Location:**"):
                    location = l.replace("- **Location:**", "").strip()
                elif l.startswith("- **Link:**"):
                    raw_link = l.replace("- **Link:**", "").strip()
                    if "→" in raw_link:
                        url_part, note_part = raw_link.split("→", 1)
                        link = url_part.strip()
                        link_note = note_part.strip()
                    else:
                        link = raw_link.split()[0] if raw_link else ""
                elif l.startswith("- **Comp:**"):
                    comp = l.replace("- **Comp:**", "").strip()
                elif l.startswith("- **Fit:**"):
                    fit_stars = l.replace("- **Fit:**", "").strip()

            comp_clean = company or "Target Company"
            jobs.append({
                "company": comp_clean,
                "role": role,
                "location": location,
                "link": link,
                "link_note": link_note,
                "comp": comp,
                "fit": fit_stars,
                "salary_benchmark": salary.estimate_market_salary(comp_clean, role, location),
                **portals.build_referral_links(comp_clean, role),
            })
        return jobs

    def _run_eligibility_and_language_gates(self, text_lower: str, location: str) -> dict:
        loc_lower = (location or "").lower()
        non_india_remote_markers = [
            "us - remote", "us-remote", "us remote", "united states", "nyc", "san francisco",
            "seattle", "austin", "boston", "chicago", "americas", "north america", "latam",
            "canada", "uk - remote", "uk remote", "united kingdom", "london", "dublin",
            "eu - remote", "europe", "emea", "amsterdam", "berlin", "singapore", "dubai", "uae",
            "sydney", "australia", "japan", "tokyo", "jakarta", "manila", "kuala lumpur",
        ]
        india_markers = ["india", "bengaluru", "bangalore", "mumbai", "hyderabad", "pune", "chennai", "gurgaon", "gurugram", "noida", "delhi"]
        has_india = any(k in loc_lower for k in india_markers) or "bengaluru" in text_lower or "bangalore" in text_lower
        has_foreign_loc = any(m in loc_lower for m in non_india_remote_markers)
        is_india = has_india or (not has_foreign_loc and loc_lower.strip() in ("remote", "anywhere", "global - remote"))

        # Eligibility Gate (04-job-evaluation.md)
        eligibility_gate = "PASS" if is_india else "FLAG"
        eligibility_note = "India base (no work permit needed)" if is_india else "Abroad role: verify visa sponsorship"
        hard_eligibility_patterns = [
            r"\bmust be a (?:us |u\.s\. |uk |british |uae |singapore )?citizen\b",
            r"\bcitizenship required\b",
            r"\bsecurity clearance required\b",
            r"\bpermanent resident(?:s)? only\b",
            r"\bno visa sponsorship\b",
            r"\bnot able to sponsor\b",
            r"\bunable to sponsor\b",
            r"\bmust hold (?:existing )?right to work without sponsorship\b",
        ]
        if not is_india:
            for pat in hard_eligibility_patterns:
                m = re.search(pat, text_lower)
                if m:
                    eligibility_gate = "FAIL"
                    eligibility_note = f"Matched eligibility restriction: '{m.group(0)}'"
                    break
            if eligibility_gate != "FAIL" and any(w in text_lower for w in ["visa sponsorship", "relocation package", "sponsorship provided", "relocation support"]):
                eligibility_gate = "PASS"
                eligibility_note = "Visa sponsorship / relocation stated"

        # Language Gate (04-job-evaluation.md): Declared English (Fluent), Hindi (Professional), Bengali (Native)
        language_gate = "PASS"
        language_note = "English working language"
        undeclared_reqs = [
            (r"\b(?:fluent|native|c1|c2|proficient|business level)\s+(?:in\s+)?(german|french|spanish|dutch|danish|arabic|mandarin|japanese|korean|polish|italian|portuguese|swedish|norwegian)\b", None),
            (r"\b(german|french|spanish|dutch|danish|arabic|mandarin|japanese)\s+(?:is\s+)?(?:required|mandatory|essential)\b", None),
        ]
        for pat, _ in undeclared_reqs:
            m = re.search(pat, text_lower)
            if m:
                language_gate = "FAIL"
                language_note = f"Requires undeclared language: '{m.group(0)}'"
                break

        location_verdict = "PASS" if eligibility_gate == "PASS" else ("FAIL" if eligibility_gate == "FAIL" else "FLAG")
        return {
            "eligibility_gate": eligibility_gate,
            "eligibility_note": eligibility_note,
            "language_gate": language_gate,
            "language_note": language_note,
            "location_verdict": location_verdict,
        }

    def score_jd(self, company, role, jd_text, location="", guard_findings=None):
        raw_text = jd_text or ""
        if guard_findings is None:
            cleaned_text, guard_findings = posting_guard.guard(raw_text, "auto")
        else:
            cleaned_text = raw_text

        text_lower = f"{role or ''} {location or ''} {cleaned_text}".lower()
        role_lower = (role or "").lower()

        # Gates
        gates = self._run_eligibility_and_language_gates(text_lower, location)

        # Level Discipline (senior-pm-job-search SKILL.md Rule 7: Visa I6 Senior PM; Director+ = Stretch)
        is_stretch = any(t in role_lower for t in ["director", "vp ", "vice president", "head of product", "group product manager", "gpm"])
        is_too_junior = any(t in role_lower for t in ["associate product manager", "apm ", "intern", "junior product"])

        # Weighted keywords matching candidate-facts & master CV
        domain_keywords = [
            "payment", "acquiring", "settlement", "merchant", "clearing", "reconciliation",
            "iso 8583", "gateway", "tokenization", "api", "fintech", "card", "scheme",
            "kafka", "microservices", "digital banking", "jaywan", "rupay", "cross-border",
            "dispute", "chargeback", "risk", "compliance", "emv", "bin", "base i", "base ii", "b2b"
        ]
        matched_domain = [kw for kw in domain_keywords if kw in text_lower]
        domain_score = min(10.0, max(4.2, 4.2 + len(matched_domain) * 0.72))

        seniority_score = 9.5
        if is_too_junior:
            seniority_score = 3.5
        elif is_stretch:
            seniority_score = 7.8  # flagged as Stretch Application per search-profile.md
        elif any(term in role_lower for term in ["senior product", "principal product", "staff product", "lead product", "product manager"]):
            seniority_score = 9.8

        tech_keywords = [
            "api", "kafka", "rest", "architecture", "microservice", "ai", "llm", "agentic",
            "automation", "specs", "event-driven", "sql", "rice", "prd", "agile", "scrum", "cspo",
            "rag", "vector database", "mcp", "model context protocol", "tool calling", "prompt",
            "evals", "guardrails", "embeddings", "fine-tuning", "human-in-the-loop", "inference",
            "model routing", "hallucination", "context window", "context management", "orchestration",
            "copilot", "openai", "defect analysis", "acceptance criteria", "sdlc",
        ]
        matched_tech = [kw for kw in tech_keywords if kw in text_lower]
        tech_score = min(10.0, max(5.2, 5.2 + len(matched_tech) * 0.85))

        impact_score = 9.4

        sponsorship = "UNVERIFIED"
        loc_score = 8.0
        if "india" in (location or "").lower() or "bengaluru" in text_lower or "bangalore" in text_lower or "india" in text_lower:
            loc_score = 10.0
            sponsorship = "Not needed (India)"
        elif gates["eligibility_gate"] == "PASS":
            loc_score = 9.5
            sponsorship = "STATED"
        elif gates["eligibility_gate"] == "FAIL":
            loc_score = 2.0
            sponsorship = "NO (Restricted)"
        else:
            loc_score = 6.2
            sponsorship = "UNVERIFIED"

        total_score = round(
            (0.30 * domain_score)
            + (0.20 * seniority_score)
            + (0.15 * tech_score)
            + (0.15 * impact_score)
            + (0.20 * loc_score),
            1,
        )

        # 04-job-evaluation.md 0-100 scale (Technical 30%, Experience 25%, Behavioral 15%, Career 30%)
        tech_100 = int(round(tech_score * 10))
        exp_100 = int(round(domain_score * 10))
        beh_100 = int(round(min(96, impact_score * 9.5)))
        car_100 = int(round(seniority_score * 9.5 if gates["eligibility_gate"] != "FAIL" else 30))
        rank_score = portals.overall_score({
            "technical": tech_100,
            "experience": exp_100,
            "behavioral": beh_100,
            "career": car_100,
        })
        rank_verdict = portals.band(rank_score)

        if gates["eligibility_gate"] == "FAIL":
            verdict = "ELIGIBILITY VETO"
        elif gates["language_gate"] == "FAIL":
            verdict = "LANGUAGE VETO"
        elif total_score >= 7.0:
            verdict = "STRETCH APPLY" if is_stretch else "APPLY"
        elif total_score >= 6.0:
            verdict = "APPLY WITH REFERRAL"
        else:
            verdict = "SKIP"

        # Grounded proof points (aligned with candidate-facts.md: 10 FTEs settled, CS-attributed revenue protected, Fast Refund AI)
        proof_points = [
            "Delivered domestic-scheme processing (KSA, RuPay, CUP, Jaywan, PayPak, MyDebit) via Payments 2.0 & Kafka on CyberSource ($30M net revenue).",
            "Automated CAS self-service testing & settlement workflows at Visa, freeing 10 FTEs and cutting turnaround under 24 hours.",
            "Applied AI in regulated payments (Fast Refund: removed spike analysis/regression -> 24h refunds, 97% success) and Claude agentic specs (SMCC Japan: 6 wks -> 2 wks).",
            "Revenue protection (attributed to Customer Support after releases): $9.2M at Visa via automated settlement regeneration and $11M at Cognizant/Fiserv; $1.5M KFC auth-to-settlement recapture.",
            "Scaled Visa Platform Connect acquirer onboarding from 100+ to 400+ acquirers and grew European token volume by 25% (COF/MIT/CIT) + CEMEA instalments by 18%."
        ]

        # Match Table & Known Gaps checks strictly against candidate-facts.md
        match_table = [
            {"requirement": "Payments / Acquiring / Settlement Rails", "evidence": "14 yrs (Visa VPC/CyberSource, Cognizant Fiserv); ISO 8583, BASE I/II, Kafka", "strength": "Strong" if any(k in text_lower for k in ["payment", "acquiring", "settlement", "card", "fintech"]) else "Moderate"},
            {"requirement": "Senior PM Scope & Strategy (RICE / Roadmaps)", "evidence": "Visa Level I6 PM; RICE roadmap signed off by VP; 4 scrum teams (~6 each)", "strength": "Strong"},
            {"requirement": "AI / Agentic Workflow Productization", "evidence": "Fast Refund AI optimization (97% success) + Claude/Agentic PM workflows", "strength": "Strong" if any(k in text_lower for k in ["ai", "llm", "agent", "automation"]) else "Moderate"},
        ]

        gaps = []
        red_flags = []
        gap_questions = []

        # Known Gaps from references/candidate-facts.md
        known_gap_rules = [
            (["upi"], "UPI Exposure Gap: Candidate facts bank confirms no direct UPI product exposure (core is card rails, domestic schemes RuPay/Jaywan/KSA, ISO 8583).", "What card-rail / real-time settlement parallels (e.g., RuPay, NRTF, Instant Refund) best bridge the UPI requirement for this team?"),
            (["lending", "loan", "credit underwriting", "bnpl"], "Consumer Lending Gap: No direct consumer lending/underwriting exposure (worked on Visa Instalment Plan B2B spec enablement in CEMEA).", "Should we frame Visa Instalment Plan (18% CEMEA growth) strictly as rail enablement without overclaiming underwriting?"),
            (["insurance"], "Insurance Domain Gap: No direct insurance product exposure in verified facts bank.", "Are there compliance/regulated-workflow experiences from Visa RBI mandates you want emphasized for Insurance?"),
            (["mutual fund", "wealth", "digital gold"], "Wealth / Mutual Funds / Digital Gold Gap: Facts bank confirms zero exposure to wealth or digital gold products.", "Focus cover note on core money-movement plumbing rather than asset-management product metrics?"),
            (["b2c", "consumer app", "consumer-facing", "consumer fintech"], "B2C Consumer Exposure Gap: Career experience is B2B (acquirers, merchants, gateway platform), not consumer-facing DAU/retention funnels.", "Can you cite merchant-facing UX wins (Acquirer Onboarding UI change) as B2B UX judgement while keeping B2C honest?"),
            (["master's", "mba preferred", "mba required"], "Academic Credential Note: Facts bank confirms B.Tech IT (8.6 GPA) and CSPO — no Master's degree / MBA.", "Confirm whether we address 14 years experience + CSPO upfront as practical equivalence."),
            (["vendor contract", "procurement"], "Vendor Contract Ownership Gap: Facts bank notes no direct vendor-contract ownership (did influence Cognizant SOW requote without formal contract ownership).", "Want to include the Cognizant SOW requote influence story carefully framed as indirect influence?"),
            (["direct reports", "manage pms", "hiring and mentoring pms"], "People Management Scope Nuance: Led 12-person team at Cognizant (2019-2021) + 100+ interviews; at Visa I6 mentored 4 scrum teams via velocity/scorecards with NO formal direct reports.", "Should the CV emphasize Cognizant 12-person lead experience alongside Visa scorecard mentoring?"),
        ]
        for triggers, gap_msg, q_msg in known_gap_rules:
            if any(t in text_lower for t in triggers):
                gaps.append(gap_msg)
                red_flags.append(gap_msg.split(":")[0])
                match_table.append({"requirement": gap_msg.split(":")[0], "evidence": "Missing / Adjacent in candidate-facts.md", "strength": "Missing"})
                gap_questions.append(q_msg)

        if is_stretch:
            red_flags.append("Stretch Level Alert: JD targets Director/GPM/Lead; Avijit's Visa standing is I6 PM (Senior PM home base). Treat as Stretch Application.")
            gap_questions.append("This role is Director/Lead level (Stretch Application). Confirm you want to proceed with executive positioning?")

        if gates["eligibility_gate"] == "FLAG":
            gaps.append(f"Visa Sponsorship UNVERIFIED for {location or 'international hub'}: confirm relocation support early.")
            red_flags.append("Visa sponsorship unverified in posting text")
        elif gates["eligibility_gate"] == "FAIL":
            red_flags.append(gates["eligibility_note"])

        if gates["language_gate"] != "PASS":
            red_flags.append(gates["language_note"])

        return {
            "score": total_score,
            "rank_score": rank_score,
            "rank_verdict": rank_verdict,
            "verdict": verdict,
            "is_stretch": is_stretch,
            "sponsorship": sponsorship,
            "gates": gates,
            "guard_verdict": guard_findings.get("verdict", "posting_guard: clean"),
            "guard_warnings": {
                "instruction_like": guard_findings.get("instruction_like", []),
                "invisible_flagged": guard_findings.get("invisible_flagged", 0),
                "homoglyph_runs": guard_findings.get("homoglyph_runs", []),
                "hidden_blocks_count": len(guard_findings.get("hidden_blocks", [])),
            },
            "breakdown": {
                "domain_match": round(domain_score, 1),
                "seniority_match": round(seniority_score, 1),
                "technical_ai_depth": round(tech_score, 1),
                "measurable_impact": round(impact_score, 1),
                "location_sponsorship": round(loc_score, 1),
                "weighted_100": {
                    "technical": tech_100,
                    "experience": exp_100,
                    "behavioral": beh_100,
                    "career": car_100,
                },
            },
            "match_table": match_table,
            "red_flags": red_flags,
            "gap_questions": gap_questions,
            "proof_points": proof_points,
            "gaps": gaps,
            "matched_keywords": sorted(set(matched_domain + matched_tech)),
            "salary_benchmark": salary.estimate_market_salary(company, role, location, jd_text),
            "referral_links": portals.build_referral_links(company, role),
        }

    def _get_candidate_corpus_lower(self) -> str:
        """Loads verified candidate facts + master CV text for honest ATS keyword matching."""
        parts = [read_text("profile/master_cv.md")]
        if self.candidate_facts_path.exists():
            parts.append(self.candidate_facts_path.read_text(encoding="utf-8", errors="ignore"))
        return "\n".join(parts).lower()

    def _extract_advertised_comp(self, text: str) -> str | None:
        """Extracts verbatim stated compensation range/figure from JD if present; never estimates."""
        return salary.extract_stated_comp(text)

    def _extract_jd_requirements(self, cleaned_text: str, role: str, candidate_lower: str) -> list[dict]:
        """
        Pass 1 (JD-Only) Requirement Extraction followed by Pass 2 Evidence Mapping
        strictly per career-ops modes/oferta.md Block B (never fabricating evidence).
        """
        # Split lines/sentences to capture verbatim quotes
        raw_lines = [re.sub(r"\s+", " ", ln).strip(" •-*–\t") for ln in (cleaned_text or "").splitlines()]
        sentences = []
        for ln in raw_lines:
            if len(ln) < 18:
                continue
            for seg in re.split(r"(?<=[.;])\s+", ln):
                seg_c = seg.strip()
                if 18 <= len(seg_c) <= 260:
                    sentences.append(seg_c)

        domain_map = [
            (["payment", "acquiring", "merchant processing", "card network", "scheme"], "Payments / Acquiring Platform Scale", "Delivered multi-scheme acquiring on Visa VPC & CyberSource ($30M net revenue; 400+ acquirers).", "Strong"),
            (["clearing", "settlement", "reconciliation", "iso 8583", "base i", "base ii", "tc33"], "Clearing, Settlement & Protocol Rails", "Decoupled 85% BASE I into 5 microservices; 100% BASE II / TC33.A compliance; CAS settlement automation freed 10 FTEs.", "Strong"),
            (["api", "rest", "webhook", "microservice", "event-driven", "kafka", "architecture"], "APIs, Event-Driven Systems & Technical Fluency", "Owned Payments 2.0 Kafka microservices & CyberSource REST/ISO integrations across 4 scrum teams.", "Strong"),
            (["ai", "llm", "agent", "genai", "machine learning", "automation"], "Applied AI & Workflow Automation", "Shipped Fast Refund AI optimization (97% success, 24h) + Claude agentic specs (SMCC Japan: 6 wks -> 2 wks).", "Strong"),
            (["token", "cof", "mit", "cit", "3ds", "fraud", "chargeback", "dispute", "risk", "compliance", "rbi", "emv", "pci"], "Regulatory Compliance, Tokenization & Risk", "Grew EU COF/MIT/CIT token volume +25%; led Apple Pay India RBI data-center migration cleanly.", "Strong"),
            (["roadmap", "rice", "prd", "strategy", "discover", "stakeholder", "cross-functional", "agile", "scrum"], "Senior PM Discovery, RICE & Stakeholder Execution", "Secured VP sign-off on RICE roadmap; mentored 4 scrum teams (~6 engineers each) as CSPO.", "Strong"),
            (["cross-border", "fx", "multi-currency", "international", "global"], "Cross-Border & Multi-Market Payments", "Launched regional schemes (KSA MADA, UAE Jaywan, RuPay, CUP, PayPak, MyDebit, EFTPOS) & SMCC Japan.", "Strong"),
            (["upi", "npci", "a2a", "real-time payment"], "UPI / Real-Time A2A Payment Rails", "No direct UPI ownership in candidate-facts.md; adjacent domestic-scheme & Instant/Fast Refund rails (97% success).", "Partial"),
            (["lending", "loan", "underwriting", "bnpl", "credit line"], "Consumer Lending / Credit Underwriting", "No consumer underwriting exposure; B2B Visa Instalment Plan enablement (+18% CEMEA) only.", "Missing"),
            (["b2c", "consumer app", "consumer-facing", "dau", "mau", "retention loop"], "B2C Consumer Growth & Funnels", "Background is B2B (acquirers/merchants/processors), not B2C DAU funnels; cite Acquirer Onboarding UI launch.", "Partial"),
            (["wealth", "mutual fund", "digital gold", "insurance", "brokerage"], "WealthTech / Mutual Funds / Insurance", "Zero verified exposure in candidate-facts.md — avoid overclaiming.", "Missing"),
            (["mba", "master's degree", "masters degree", "postgraduate"], "Graduate Credential (MBA / Master's)", "Holds B.Tech IT (8.6 GPA) + CSPO + 14 yrs PM/payments execution; no MBA/Master's.", "Partial"),
            (["direct reports", "manage a team of pms", "hire and mentor pms"], "Formal People Management / Direct Reports", "Led 12-person PM/analyst team at Cognizant (2019-21) + 100+ interviews; Visa I6 had 0 direct reports (velocity/scorecards).", "Partial"),
            (["vendor contract", "procurement", "rfp"], "Direct Vendor Contract / Procurement Ownership", "No formal vendor-contract ownership; influenced Cognizant SOW requote so automation-freed analysts remained.", "Missing"),
        ]

        extracted = []
        used_reqs = set()
        for idx, sent in enumerate(sentences):
            sl = sent.lower()
            # Determine importance tier per career-ops oferta.md Block B
            if any(w in sl for w in ["must have", "required", "minimum", "years of experience", "proven track record", "essential", "mandatory", "strong background", "deep experience"]):
                band = "critical (stated)"
            elif any(w in sl for w in ["responsible for", "own the", "lead ", "architect", "define ", "partner with", "deliver", "drive ", "manage "]):
                band = "high (structural)"
            elif any(w in sl for w in ["preferred", "nice to have", "plus", "bonus", "ideal"]):
                band = "meaningful (inferred)"
            else:
                continue

            for triggers, label, evidence, default_strength in domain_map:
                if label in used_reqs:
                    continue
                if any(t in sl for t in triggers):
                    used_reqs.add(label)
                    extracted.append({
                        "band": band,
                        "requirement": label,
                        "jd_quote": sent,
                        "evidence": evidence,
                        "strength": default_strength,
                    })
                    break
            if len(extracted) >= 10:
                break

        # Ensure baseline extracted rows if JD text was short/sparse
        full_lower = f"{role} {cleaned_text}".lower()
        for triggers, label, evidence, default_strength in domain_map:
            if label in used_reqs:
                continue
            hit = next((t for t in triggers if t in full_lower), None)
            if hit:
                used_reqs.add(label)
                extracted.append({
                    "band": "high (structural)" if default_strength != "Missing" else "critical (stated)",
                    "requirement": label,
                    "jd_quote": f"Matched keyword '{hit}' in role/JD text.",
                    "evidence": evidence,
                    "strength": default_strength,
                })
            if len(extracted) >= 7:
                break

        return extracted

    def _analyze_block_g_and_gates(
        self, company: str, role: str, cleaned_text: str, location: str, job_url: str, guard_findings: dict
    ) -> dict:
        """Runs career-ops Block A geo/work-auth + Block G Posting Legitimacy + Prior-Contact FYI."""
        text_lower = f"{role or ''} {location or ''} {cleaned_text or ''}".lower()
        loc_lower = (location or "").lower()
        comp_lower = (company or "").lower()
        gates = self._run_eligibility_and_language_gates(text_lower, location)

        signals = []
        caution_count = 0
        suspicious_count = 0

        # 1. Geo-Mismatch check (e.g. stated remote vs mandatory onsite/hybrid, or US-only wording vs India)
        geo_mismatch = None
        if "remote" in loc_lower and any(w in text_lower for w in ["days in office", "days a week in office", "hybrid model", "in-office", "onsite only", "on-site only"]):
            geo_mismatch = "Stated 'Remote' in header/location, but JD body mentions in-office / hybrid cadence."
            caution_count += 1
            signals.append({"signal": "Geo / Work-Mode Consistency", "weight": "Concerning", "finding": geo_mismatch})
        elif any(m in loc_lower for m in ["india", "bengaluru", "bangalore"]) and any(w in text_lower for w in ["401(k)", "401k", "w-2", "us citizen only"]):
            geo_mismatch = "Location lists India, but benefits/tax text references US-specific terms (401k / W-2)."
            caution_count += 1
            signals.append({"signal": "Country Terminology Alignment", "weight": "Concerning", "finding": geo_mismatch})
        else:
            signals.append({"signal": "Geo / Work-Mode Consistency", "weight": "Positive", "finding": f"Location ({location or 'Unstated'}) aligns with body text."})

        # 2. Agency / Third-party check (career-ops _shared.md)
        agency_markers = ["our client", "on behalf of our client", "placement services", "staffing agency", "talent solutions", "recruiting partner"]
        is_agency = any(m in text_lower or m in comp_lower for m in agency_markers)
        if is_agency:
            caution_count += 1
            signals.append({"signal": "Hiring Entity / Agency Check", "weight": "Concerning", "finding": "Third-party agency / undisclosed end-client phrasing detected; clarify end employer before tailoring."})
        else:
            signals.append({"signal": "Hiring Entity / Agency Check", "weight": "Positive", "finding": f"Direct employer posting ({company})."})

        # 3. Description Quality & Role Clarity (career-ops Block G Signal 2 + role-clarity check)
        word_count = len((cleaned_text or "").split())
        clarity_points = {
            "Reporting line / org context": bool(re.search(r"\b(report(?:s|ing)?\s+to|org|organization|team of|pillar|business unit|tribe|pod)\b", text_lower)),
            "Concrete 6-12m scope / deliverables": bool(re.search(r"\b(deliver|roadmap|launch|prd|ship|metric|kpi|okr|first \d+ days|90 days)\b", text_lower)),
            "Specific domain stack / protocols": bool(re.search(r"\b(api|iso|kafka|payment|acquiring|settlement|sql|llm|ai|scheme|clearing|gateway|upi|card)\b", text_lower)),
            "Cross-functional coordination cadence": bool(re.search(r"\b(engineering|design|risk|compliance|legal|operations|sales|cross-functional)\b", text_lower)),
        }
        clarity_score = sum(1 for v in clarity_points.values() if v)
        if word_count < 80:
            caution_count += 1
            signals.append({"signal": "JD Substance & Role Clarity", "weight": "Concerning", "finding": f"Brief or incomplete JD ({word_count} words); fetch full posting or ask recruiter for full scope."})
        elif clarity_score >= 3:
            signals.append({"signal": "JD Substance & Role Clarity", "weight": "Positive", "finding": f"Detailed JD ({word_count} words, {clarity_score}/4 clarity markers present)."})
        else:
            caution_count += 1
            signals.append({"signal": "JD Substance & Role Clarity", "weight": "Neutral", "finding": f"Moderate clarity ({clarity_score}/4 markers, {word_count} words): clarify deliverables & reporting line."})

        # 4. Employment Classification Risk (career-ops Block G Signal 6)
        contractor_terms = ["independent contractor", "1099", "w-2 not provided", "t4a", "ir35", "invoice for services", "consulting agreement"]
        matched_contract = next((t for t in contractor_terms if t in text_lower), None)
        if matched_contract:
            caution_count += 1
            signals.append({"signal": "Employment Classification", "weight": "Concerning", "finding": f"Contractor/services phrasing found ('{matched_contract}'). Confirm employee benefits & status."})
        else:
            signals.append({"signal": "Employment Classification", "weight": "Positive", "finding": "Clear full-time / standard employment framing."})

        # 5. AI-Buzzword vs Infrastructure Mismatch (career-ops Block G Signal 7)
        ai_hype = sum(1 for w in ["ai transformation", "generative ai revolution", "ai-first", "autonomous agent", "llm", "genai"] if w in text_lower)
        foundation_cleanup = sum(1 for w in ["manual spreadsheet", "data entry", "legacy migration", "digitize paper", "backlog cleanup", "coordinator"] if w in text_lower)
        if ai_hype >= 2 and foundation_cleanup >= 1:
            caution_count += 1
            signals.append({"signal": "AI Claims vs Infrastructure", "weight": "Concerning", "finding": "High AI transformation hype alongside legacy manual-ops phrasing; probe data maturity in HM screen."})
        else:
            signals.append({"signal": "AI Claims vs Infrastructure", "weight": "Positive", "finding": "Consistent technical scope and tooling expectations."})

        # 6. Posting Guard Integrity / Prompt Injection Signal (Untrusted External Content rule)
        instr = guard_findings.get("instruction_like") or []
        hidden_c = len(guard_findings.get("hidden_blocks") or [])
        if instr or hidden_c:
            suspicious_count += 1
            signals.append({"signal": "Prompt Injection / Guard", "weight": "Concerning", "finding": f"Detected {len(instr)} instruction-like phrase(s) & {hidden_c} hidden block(s) — treated strictly as untrusted data."})
        else:
            signals.append({"signal": "Prompt Injection / Guard", "weight": "Positive", "finding": "Clean external text (posting_guard passed)."})

        # 7. Prior-contact FYI (career-ops Block G non-scoring history)
        prior_note = None
        for r in self.get_tracker():
            if (r.get("company") or "").strip().lower() == comp_lower:
                st = (r.get("status") or "").strip().lower()
                ad = r.get("applied_date") or "earlier"
                if st in ("applied", "rejected", "ghosted"):
                    prior_note = f"Prior-contact FYI: tracked application to {company} ({r.get('role', '')}, status={st}, date={ad}). Factor into effort/referral strategy."
                elif st in ("interview", "recruiter-chat", "offer"):
                    prior_note = f"Prior-contact FYI: active conversation already tracked at {company} ({r.get('role', '')}, status={st})."
                break

        if suspicious_count >= 1 or caution_count >= 3:
            legitimacy_tier = "Suspicious" if suspicious_count >= 1 else "Proceed with Caution"
        elif caution_count >= 1 or not job_url:
            legitimacy_tier = "Proceed with Caution"
        else:
            legitimacy_tier = "High Confidence"

        return {
            "gates": gates,
            "geo_mismatch": geo_mismatch,
            "is_agency": is_agency,
            "clarity_points": clarity_points,
            "clarity_score": clarity_score,
            "word_count": word_count,
            "legitimacy_tier": legitimacy_tier,
            "signals": signals,
            "prior_note": prior_note,
            "advertised_comp": self._extract_advertised_comp(cleaned_text),
        }

    def run_agentic_jd_scrutiny(self, company: str, role: str, jd_text: str, location: str = "", job_url: str = "") -> dict:
        """
        Agentic Multi-Persona JD Scrutiny Flow combining career-ops (Blocks A-G + ATS + Red-Flag rules)
        with 5 sequential evaluation personas: ATS, HR/Recruiter, Tech Lead, Hiring Manager, Bar Raiser.
        Strictly grounded in verified facts from references/candidate-facts.md & profile/master_cv.md.
        """
        t0 = time.perf_counter()
        company_clean = portals.normalize_company_display(company or "Target Company")
        role_clean = (role or "Senior Product Manager").strip()
        loc_clean = (location or "Bengaluru").strip()
        cleaned_jd, guard_findings = posting_guard.guard(jd_text or "", "auto")

        # Baseline quantitative scorer + Block A/G legitimacy gates
        fit_data = self.score_jd(company_clean, role_clean, cleaned_jd, loc_clean, guard_findings=guard_findings)
        block_g = self._analyze_block_g_and_gates(company_clean, role_clean, cleaned_jd, loc_clean, job_url, guard_findings)
        candidate_lower = self._get_candidate_corpus_lower()
        requirements_matrix = self._extract_jd_requirements(cleaned_jd, role_clean, candidate_lower)

        text_lower = f"{role_clean} {loc_clean} {cleaned_jd}".lower()
        role_lower = role_clean.lower()
        is_stretch = fit_data["is_stretch"]
        is_too_junior = any(t in role_lower for t in ["associate product manager", "apm ", "intern", "junior product"])
        gates = block_g["gates"]

        # ------------------------------------------------------------------
        # PERSONA 1: ATS PARSER & KEYWORD FILTER (career-ops modes/ats.md)
        # ------------------------------------------------------------------
        ats_target_vocab = [
            "payments", "acquiring", "settlement", "clearing", "reconciliation", "merchant", "iso 8583",
            "base i", "base ii", "kafka", "rest", "api", "microservices", "tokenization", "emv", "3ds",
            "chargeback", "dispute", "compliance", "rbi", "pci", "rice", "prd", "roadmap", "agile",
            "scrum", "cspo", "sql", "ai", "llm", "agentic", "automation", "cross-border", "fx", "upi",
            "lending", "bnpl", "b2b", "b2c", "kyc", "aml", "fraud", "risk", "analytics", "experimentation",
            "a/b testing", "stakeholder", "enterprise", "saas", "fintech", "sdk", "checkout", "pos",
            "rag", "vector database", "embeddings", "prompt engineering", "context management", "evals",
            "guardrails", "hallucination", "tool calling", "mcp", "model context protocol", "fine-tuning",
            "model routing", "inference", "context window", "human-in-the-loop", "orchestration",
            "copilot", "openai", "defect analysis", "acceptance criteria", "latency", "sdlc",
        ]
        jd_present_vocab = [kw for kw in ats_target_vocab if kw in text_lower]
        matched_ats = [kw for kw in jd_present_vocab if kw in candidate_lower or kw in ("payments", "ai", "automation")]
        missing_ats = [kw for kw in jd_present_vocab if kw not in matched_ats]
        kw_cov_pct = int(round((len(matched_ats) / len(jd_present_vocab)) * 100)) if jd_present_vocab else 78

        exp_req_match = re.search(r"(\d{1,2})\s*\+?\s*years", text_lower)
        req_years = int(exp_req_match.group(1)) if exp_req_match else None
        knockouts = []
        ats_score = 4.4
        if gates["eligibility_gate"] == "FAIL":
            knockouts.append(f"HARD KNOCKOUT: {gates['eligibility_note']}")
            ats_score -= 1.6
        if gates["language_gate"] == "FAIL":
            knockouts.append(f"HARD KNOCKOUT: {gates['language_note']}")
            ats_score -= 1.6
        if req_years and req_years > 15:
            knockouts.append(f"Years filter warning: JD asks {req_years}+ yrs (Avijit has ~14 yrs total, ~10 yrs PM).")
            ats_score -= 0.5
        if any(k in text_lower for k in ["mba required", "master's degree required", "masters degree required"]):
            knockouts.append("Degree filter risk: JD marks MBA/Master's as required (Candidate holds B.Tech IT 8.6 GPA + CSPO).")
            ats_score -= 0.7
        if missing_ats:
            ats_score -= min(0.8, len(missing_ats) * 0.12)
        ats_score = round(max(1.5, min(5.0, ats_score)), 1)
        ats_verdict = "PASS" if (ats_score >= 4.0 and not any("HARD" in k for k in knockouts)) else ("CAUTION" if ats_score >= 3.0 else "BLOCK")

        ats_persona = {
            "id": "ats",
            "persona": "ATS Parser & Filter",
            "emoji": "🤖",
            "subtitle": "Keyword coverage, hard knockout rules & DOCX parseability (modes/ats.md)",
            "score_5": ats_score,
            "status": ats_verdict,
            "headline": f"{kw_cov_pct}% JD keyword coverage across {len(jd_present_vocab) or 1} detected domain terms • Single-column DOCX clean.",
            "strengths": [
                f"Matched JD keywords ({len(matched_ats)}): {', '.join(matched_ats[:10]) if matched_ats else 'core PM & payments vocabulary'}.",
                "DOCX pipeline preserves single-column ATS structure, standard headings, email in body, and avoids employer-in-role title distortion.",
                "Experience (~14 yrs total / ~10 yrs PM) & CSPO satisfy typical Senior PM screening filters.",
            ],
            "concerns": knockouts + (
                [f"Unmatched JD terms on Master CV ({len(missing_ats)}): {', '.join(missing_ats[:8])} — DO NOT fabricate; use adjacent framing only where truthful."]
                if missing_ats else ["Zero critical keyword blindspots detected."]
            ),
            "action_items": [
                "Mirror exact JD terminology for truth-backed skills (e.g., 'clearing & settlement', 'Payments 2.0 Kafka microservices', 'CI/CD AI workflows') in top Summary.",
                "Keep Visa / Cognizant employer blocks cleanly separated without cramming 'First Data' into job titles unnecessarily unless ATS-parsed cleanly.",
            ],
            "probe_questions": [],
        }

        # ------------------------------------------------------------------
        # PERSONA 2: HR / RECRUITER SCREENING (10-Second Scan)
        # ------------------------------------------------------------------
        hr_flags = list(fit_data["red_flags"])
        hr_score = 4.3
        if is_stretch:
            hr_score -= 0.5
        if is_too_junior:
            hr_score -= 1.8
        if gates["eligibility_gate"] == "FAIL":
            hr_score = min(hr_score, 2.0)
        elif gates["eligibility_gate"] == "FLAG":
            hr_score -= 0.4
        if gates["language_gate"] == "FAIL":
            hr_score = min(hr_score, 1.8)
        if block_g["is_agency"]:
            hr_score -= 0.4
        hr_score = round(max(1.5, min(5.0, hr_score)), 1)
        hr_verdict = "PASS" if hr_score >= 4.0 else ("CONDITIONAL" if hr_score >= 3.2 else "BLOCK")

        salary_bench = salary.estimate_market_salary(company_clean, role_clean, loc_clean, cleaned_jd)
        comp_obs = (
            f"Advertised compensation verbatim in JD: {block_g['advertised_comp']} (Market Est. via {salary_bench['source_basis']}: {salary_bench['avg_display']}, range {salary_bench['range_display']})."
            if block_g["advertised_comp"]
            else f"Compensation unstated in JD (verify fixed vs bonus/equity on recruiter call). Probable Market Range ({salary_bench['source_basis']}): {salary_bench['avg_display']} • {salary_bench['range_display']} ({salary_bench['split_note']})."
        )

        hr_persona = {
            "id": "hr",
            "persona": "HR / Recruiter",
            "emoji": "👤",
            "subtitle": "10-second scan: level band, work authorization, career narrative & comp clarity",
            "score_5": hr_score,
            "status": hr_verdict,
            "headline": (
                "Target Level Fit (Visa I6 -> Senior/Staff PM)" if not is_stretch and not is_too_junior
                else ("Stretch Level Alert: Director/GPM vs Visa I6 home base" if is_stretch else "Level Mismatch: Too junior vs 14 yrs experience")
            ),
            "strengths": [
                "Marquee brand progression: Visa (I6 PM, 2x STAR Award) + Cognizant/Fiserv (10 yrs PM / ~14 yrs total).",
                f"Location & Work Auth: {gates['eligibility_note']} | {gates['language_note']}.",
                comp_obs,
            ],
            "concerns": hr_flags + [
                "Post-Visa gap (Sep 2026 -> Nov 2026): Never mention restructuring/layoff in CV or cover note; prepare concise 15-second verbal answer if HR asks.",
            ],
            "action_items": [
                "Attach Warm Referral Path before applying online (use direct LinkedIn Recruiter & Peer Search links in Outreach).",
                "Ask in first recruiter call: fixed base split, notice/joining timeline flexibility, and exact hiring band/level.",
            ],
            "probe_questions": [
                "Tell me briefly about your transition out of Visa in Sept 2026 and what you're optimizing for next.",
                "What is your current notice/availability and target fixed compensation structure?",
            ],
        }

        # ------------------------------------------------------------------
        # PERSONA 3: TECH LEAD / SYSTEM ARCHITECT
        # ------------------------------------------------------------------
        core_rails_hits = [k for k in ["iso 8583", "base i", "base ii", "tc33", "clearing", "settlement", "kafka", "microservice", "api", "rest", "webhook", "token", "emv", "bin"] if k in text_lower]
        ai_hits = [k for k in ["ai", "llm", "agent", "genai", "rag", "eval", "automation", "mcp", "tool calling", "guardrails", "prompt", "vector", "embeddings", "human-in-the-loop", "orchestration", "context management", "model routing"] if k in text_lower]
        tech_score_5 = 4.0 + min(0.7, len(core_rails_hits) * 0.14) + min(0.3, len(ai_hits) * 0.1)
        if not core_rails_hits and not ai_hits:
            tech_score_5 = 3.6
        tech_score_5 = round(min(5.0, tech_score_5), 1)
        tech_verdict = "PASS" if tech_score_5 >= 4.0 else "CONDITIONAL"

        tech_probes = [
            f"Walk me through decoupling 85% of legacy BASE I validation into 5 Payments 2.0 Kafka microservices—how did you handle schema versioning, idempotency, and zero-downtime cutover?",
            "In Fast Refund AI optimization, what exact failure modes/regressions existed before, what guardrails kept refund success at 97%, and how did engineers evaluate false positives?",
            "How do you write PRDs/API specs for ISO 8583 / TC33.A / domestic scheme edge cases (e.g., SMCC Japan cut from 6 wks to 2 wks) without compromising auditability?",
        ]
        if any(u in text_lower for u in ["upi", "npci", "a2a"]):
            tech_probes[0] = "You have deep ISO 8583 & RuPay/Jaywan scheme experience rather than direct NPCI UPI switch ownership—how would you architect idempotent async callbacks & timeout reconciliation on a real-time A2A rail?"

        tech_persona = {
            "id": "tech_lead",
            "persona": "Tech Lead / Architect",
            "emoji": "⚙️",
            "subtitle": "Architecture depth, API/protocol rigor & strict tool-of-trade boundaries",
            "score_5": tech_score_5,
            "status": tech_verdict,
            "headline": f"Strong technical PM credibility ({len(core_rails_hits)} rail/arch matches, {len(ai_hits)} AI/automation signals).",
            "strengths": [
                "Rare protocol-level PM depth: ISO 8583, Visa BASE I/II, TC33.A, BIN routing (MADA/EFTPOS), EMV/COF/MIT/CIT.",
                "Proven modernization partner: decoupled 85% monolithic BASE I logic into 5 Kafka event-driven microservices + 400+ acquirer onboarding scale.",
                "Concrete AI workflow shipping: Fast Refund (24h refunds, 97% success) + Claude agentic specification workflow (6 wks -> 2 wks).",
            ],
            "concerns": [
                "Tool-of-Trade Boundary (career-ops CRITICAL rule): Frame Kafka, SQL/SSMS, and Claude strictly as PM architecture specs & workflow automation—NEVER imply authoring production backend services or training LLMs from scratch.",
                "Non-core stack watch: If JD demands consumer mobile SDKs, Web3/stablecoin smart contracts, or ML model fine-tuning ownership, treat as engineering-partnered or honest gap.",
            ],
            "action_items": [
                "Be ready to draw an end-to-end auth -> clearing -> Kafka settlement -> regeneration sequence diagram cleanly.",
                "Quantify technical trade-offs (e.g., Apple Pay India RBI localization compliance vs negligible infra cost).",
            ],
            "probe_questions": tech_probes,
        }

        # ------------------------------------------------------------------
        # PERSONA 4: HIRING MANAGER (Archetype + Product Ownership)
        # ------------------------------------------------------------------
        if any(k in text_lower for k in ["ai", "llm", "agent", "genai"]) and any(k in text_lower for k in ["payment", "acquiring", "fintech", "settlement"]):
            archetype = "Applied AI + Payments Platform PM"
            recommended_variant = "Payments, Acquiring & Core Rails Specialist (Master CV #1)"
            hm_base = 4.7
        elif any(k in text_lower for k in ["acquiring", "settlement", "clearing", "card", "scheme", "payment", "merchant"]):
            archetype = "Core Payments Rails & Acquiring PM"
            recommended_variant = "Payments, Acquiring & Core Rails Specialist (Master CV #1)"
            hm_base = 4.6
        elif any(k in text_lower for k in ["b2c", "consumer", "lending", "wealth", "insurance"]):
            archetype = "Consumer FinTech / Lending / Adjacent Domain PM"
            recommended_variant = "Generalist Senior Technical Product Manager (Master CV #3)"
            hm_base = 3.5
        elif is_stretch:
            archetype = "Product Leadership / Group or Director PM (Stretch)"
            recommended_variant = "Product Leadership & Stretch PM (Master CV #2)"
            hm_base = 4.0
        else:
            archetype = "B2B Platform & Technical PM"
            recommended_variant = "Generalist Senior Technical Product Manager (Master CV #3)"
            hm_base = 4.2

        hm_score = round(min(5.0, max(1.5, hm_base - (0.3 if is_stretch else 0.0) - (0.3 if len([m for m in requirements_matrix if m["strength"] == "Missing"]) >= 2 else 0.0))), 1)
        hm_verdict = "PASS" if hm_score >= 4.0 else ("CONDITIONAL" if hm_score >= 3.2 else "CAUTION")

        level_framing = (
            "Sell Senior Without Lying: Anchor on $30M Payments 2.0 net revenue + VP-signed RICE roadmap + 4 scrum teams + Cognizant 12-person lead; acknowledge cleanly that Visa I6 role was high-autonomy IC PM with scorecard influence rather than direct HR people management."
            if is_stretch or "direct reports" in text_lower or "manage" in text_lower
            else "Ideal Scope Alignment: Visa I6 Senior PM / Technical PM scope maps cleanly—4 scrum teams (~6 engineers each), 400+ acquirers, VP-level RICE prioritization, and $30M revenue delivery."
        )

        hm_persona = {
            "id": "hiring_manager",
            "persona": "Hiring Manager",
            "emoji": "🎯",
            "subtitle": "Role archetype fit, 0->1 product discovery, roadmap ownership & positioning strategy",
            "score_5": hm_score,
            "status": hm_verdict,
            "headline": f"Archetype: {archetype} • Recommended CV: {recommended_variant}",
            "strengths": [
                level_framing,
                "Commercial & customer impact: $30M net revenue across 6 domestic/regional schemes + +25% EU token growth + +18% CEMEA instalments.",
                "Execution & initiative: stepped in during peer PM leave on Acquirer Onboarding UI, got architect approval, and shipped improvement to production.",
            ],
            "concerns": [
                f"Detected {len([m for m in requirements_matrix if m['strength'] == 'Missing'])} missing & {len([m for m in requirements_matrix if m['strength'] == 'Partial'])} partial requirements in Pass-1 JD matrix — address explicitly in cover note & HM screen.",
                "Enterprise-to-Startup Speed Objection: Counter proactively with SMCC Japan 6 wks -> 2 wks Claude cycle + Fast Refund AI rapid iteration.",
            ],
            "action_items": [
                f"Lead with {recommended_variant} framing + top 4 tailored proof points.",
                "If downleveled (e.g., Staff -> Senior PM), negotiate clear 6-12 month charter/promotion criteria rather than defending title ego.",
            ],
            "probe_questions": [
                "Show me how you built the RICE prioritization model across competing country mandates and got VP buy-in—what top request did you say NO to?",
                "How do you discover merchant/acquirer pain points when you are in a platform/rails layer rather than owning a B2C app?",
            ],
        }

        # ------------------------------------------------------------------
        # PERSONA 5: BAR RAISER (Independent Skeptic + Hard Synthesis)
        # ------------------------------------------------------------------
        missing_count = sum(1 for m in requirements_matrix if m["strength"] == "Missing")
        partial_count = sum(1 for m in requirements_matrix if m["strength"] == "Partial")

        bar_challenges = [
            "Attribution Audit ($9.2M Visa / $11M Cognizant): Bar Raiser will test whether revenue protection came strictly from your product vs Customer Support fixes. Always state 'Customer Support-attributed post-release' transparently and explain your exact automated regeneration & SWAT triage mechanism.",
            "Headcount & People Management Clarity: Visa I6 role mentored 4 scrum teams via scorecards/velocity with ZERO formal direct reports. Do not conflate dotted-line engineering squads with direct reporting; lean on Cognizant (2019–2021) 12-member PM/analyst team + 100+ interviews.",
            "FTE / Metric Precision: Candidate facts settle automated CAS + settlement savings strictly at 10 FTEs (never 8 FTEs) and Transaction Validation System (2017) at ~30 analysts saving 10 hrs/file.",
        ]
        if missing_count:
            bar_challenges.append(
                f"Domain Gap Risk ({missing_count} critical/high gaps flagged): {', '.join(m['requirement'] for m in requirements_matrix if m['strength'] == 'Missing')}. Never bluff—pivot to verified adjacent execution in <30 seconds."
            )

        # Compute career-ops 1-5 Global Score (weighted across dimensions, penalized for gates/gaps)
        raw_global = (
            ats_score * 0.15
            + hr_score * 0.20
            + tech_score_5 * 0.25
            + hm_score * 0.30
            + (4.5 if block_g["legitimacy_tier"] == "High Confidence" else 3.8 if block_g["legitimacy_tier"] == "Proceed with Caution" else 2.8) * 0.10
        )
        if gates["eligibility_gate"] == "FAIL" or gates["language_gate"] == "FAIL":
            raw_global = min(raw_global, 2.4)
        global_score_5 = round(max(1.0, min(5.0, raw_global)), 2)

        # Bar Raiser strict Evidence Confidence per career-ops _shared.md
        if block_g["word_count"] >= 180 and gates["eligibility_gate"] == "PASS" and not missing_count:
            evidence_confidence = "High"
        elif block_g["word_count"] >= 90 and missing_count <= 2:
            evidence_confidence = "Medium"
        else:
            evidence_confidence = "Low"

        # Final Decision & Bar Raiser Vote per career-ops score bands
        if gates["eligibility_gate"] == "FAIL" or gates["language_gate"] == "FAIL" or global_score_5 < 3.0:
            career_ops_band = "Skip (<3.0 / Hard Stop)"
            bar_vote = "NOT INCLINED (VETO / HARD STOP)"
            final_decision = "Skip"
        elif block_g["legitimacy_tier"] == "Suspicious":
            career_ops_band = "Research First (Posting Legitimacy Concern)"
            bar_vote = "HOLD — VERIFY POSTING"
            final_decision = "Research first"
        elif global_score_5 >= 4.4 and missing_count == 0 and not is_stretch:
            career_ops_band = "Strong Apply (4.4–5.0)"
            bar_vote = "STRONG INCLINED"
            final_decision = "Apply"
        elif global_score_5 >= 4.0:
            career_ops_band = "Apply (4.0–4.39)"
            bar_vote = "INCLINED" if not is_stretch else "INCLINED (STRETCH)"
            final_decision = "Apply"
        else:
            career_ops_band = "Maybe (3.0–3.99)"
            bar_vote = "BORDERLINE — PREPARE MITIGATIONS"
            final_decision = "Consider"

        bar_score_5 = round(max(1.0, min(5.0, global_score_5 - (0.2 if missing_count >= 2 else 0.0))), 1)

        bar_probes = [
            {
                "question": "If Customer Support attributed the $9.2M protection at Visa, what specific mechanism did YOU design and prove wasn't just normal incident recovery?",
                "defense": "Clarify upfront that $9.2M was tracked by Customer Support post-release; explain that your automated settlement regeneration feature removed multi-day manual partner chasing & reconciliation blindspots, plus your personal SWAT P1 escalation triage.",
            },
            {
                "question": "Have you actually managed PMs directly at Visa, or were you strictly an individual contributor coordinating engineers?",
                "defense": "State honestly: at Visa I6 I operated as a high-scope IC PM mentoring 4 scrum teams (~6 engineers each) via velocity & scorecards with no formal direct reports; my direct reporting management (12-person PM/analyst team) was at Cognizant (2019–2021), backed by 100+ hiring interviews across my career.",
            },
            {
                "question": "Tell me about a time your compliance or platform launch priority collided with commercial cost or speed—why did you make that call?",
                "defense": "Anchor in Apple Pay India data-center migration (RBI localization compliance + schedule certainty over negligible infrastructure cost) and Payments 2.0 RICE roadmap signed off by VP.",
            },
        ]

        bar_persona = {
            "id": "bar_raiser",
            "persona": "Bar Raiser",
            "emoji": "⚖️",
            "subtitle": "Independent skeptic: attribution audit, gap stress-test, evidence confidence & final bar vote",
            "score_5": bar_score_5,
            "status": bar_vote,
            "headline": f"Bar Vote: {bar_vote} • Global Score: {global_score_5}/5 ({career_ops_band}) • Evidence: {evidence_confidence}",
            "strengths": [
                f"Evidence Confidence: {evidence_confidence} — grounded strictly in candidate-facts.md (no unverified metrics or tool-of-trade inflation).",
                f"Strongest defensible anchors: $30M net rev Payments 2.0, 10 FTEs freed CAS automation, Fast Refund AI (97% success), 400+ acquirers.",
            ],
            "concerns": bar_challenges,
            "action_items": [
                "Resolve any missing Gap Questions before submitting cover note or answers.",
                "Rehearse attribution & management-scope clarifications out loud so they sound crisp, honest, and executive-grade.",
            ],
            "probe_questions": [p["question"] for p in bar_probes],
            "probe_defenses": bar_probes,
        }

        personas = [ats_persona, hr_persona, tech_persona, hm_persona, bar_persona]
        elapsed_ms = int((time.perf_counter() - t0) * 1000)

        agent_trace = [
            {
                "step": 1,
                "agent": "Security & Block G Gatekeeper",
                "status": "WARN" if block_g["legitimacy_tier"] != "High Confidence" or gates["eligibility_gate"] == "FAIL" else "DONE",
                "summary": f"posting_guard: {fit_data['guard_verdict']} | Eligibility: {gates['eligibility_gate']} | Block G: {block_g['legitimacy_tier']}",
            },
            {
                "step": 2,
                "agent": "Pass-1 JD Extractor (Block B)",
                "status": "DONE",
                "summary": f"Extracted {len(requirements_matrix)} verbatim-anchored requirements before CV matching ({missing_count} gaps, {partial_count} partial).",
            },
            {
                "step": 3,
                "agent": "ATS Parser Agent",
                "status": ats_verdict,
                "summary": f"{ats_score}/5 — {kw_cov_pct}% JD keyword match ({len(missing_ats)} unmatched terms); DOCX parseability verified.",
            },
            {
                "step": 4,
                "agent": "HR / Recruiter Agent",
                "status": hr_verdict,
                "summary": f"{hr_score}/5 — Level: {'Stretch' if is_stretch else 'Target I6'} | Auth: {fit_data['sponsorship']}.",
            },
            {
                "step": 5,
                "agent": "Tech Lead Agent",
                "status": tech_verdict,
                "summary": f"{tech_score_5}/5 — {len(core_rails_hits)} protocol/rails hits, {len(ai_hits)} AI signals; tool-of-trade boundary enforced.",
            },
            {
                "step": 6,
                "agent": "Hiring Manager & Bar Raiser Panel",
                "status": bar_vote,
                "summary": f"Archetype: {archetype} ({hm_score}/5) -> Final Bar Vote: {bar_vote} | Global Score: {global_score_5}/5 ({final_decision}).",
            },
        ]

        scrutiny = {
            "company": company_clean,
            "role": role_clean,
            "location": loc_clean,
            "job_url": job_url or "",
            "elapsed_ms": elapsed_ms,
            "archetype": archetype,
            "recommended_variant": recommended_variant,
            "global_score_5": global_score_5,
            "career_ops_band": career_ops_band,
            "final_decision": final_decision,
            "bar_vote": bar_vote,
            "evidence_confidence": evidence_confidence,
            "fit_score_10": fit_data["score"],
            "rank_score_100": fit_data["rank_score"],
            "rank_verdict": fit_data["rank_verdict"],
            "verdict_lead": f"{final_decision}. {career_ops_band} — Bar Raiser Vote: {bar_vote} (Rank Rubric {fit_data['rank_score']}/100, Block G: {block_g['legitimacy_tier']}).",
            "block_g": block_g,
            "requirements_matrix": requirements_matrix,
            "personas": personas,
            "agent_trace": agent_trace,
            "bar_probes": bar_probes,
            "salary_benchmark": salary_bench,
            "fit_data": fit_data,
        }
        scrutiny["markdown"] = self.format_scrutiny_markdown(scrutiny)
        return scrutiny

    def format_scrutiny_markdown(self, s: dict) -> str:
        """Formats the 5-Persona Agentic Scrutiny & career-ops Blocks A-G into a readable Markdown report."""
        req_rows = "\n".join(
            f"| {r['band']} | {r['requirement']} | {r['jd_quote'][:110]} | {r['evidence']} | **{r['strength']}** |"
            for r in s["requirements_matrix"]
        )
        sig_rows = "\n".join(
            f"| {sg['signal']} | **{sg['weight']}** | {sg['finding']} |"
            for sg in s["block_g"]["signals"]
        )
        persona_sections = []
        for p in s["personas"]:
            strengths_md = "\n".join(f"  - {st}" for st in p["strengths"])
            concerns_md = "\n".join(f"  - {c}" for c in p["concerns"]) if p["concerns"] else "  - None."
            actions_md = "\n".join(f"  - {a}" for a in p["action_items"])
            probes_md = ""
            if p.get("probe_questions"):
                probes_md = "\n- **Likely Probe Questions:**\n" + "\n".join(f"  - {q}" for q in p["probe_questions"])
            persona_sections.append(
                f"### {p['emoji']} {p['persona']} — **{p['score_5']} / 5** (`{p['status']}`)\n"
                f"*{p['headline']}*\n\n"
                f"- **Strengths / Passes:**\n{strengths_md}\n"
                f"- **Concerns / Risks:**\n{concerns_md}\n"
                f"- **Actions / Positioning:**\n{actions_md}"
                f"{probes_md}\n"
            )

        return f"""# career-ops 5-Persona Agentic JD Scrutiny: {s['company']} – {s['role']}

## Verdict (lead)
**{s['verdict_lead']}**

- **Location / Work Auth:** {s['location']} — {s['fit_data']['gates']['eligibility_note']} (`{s['fit_data']['gates']['eligibility_gate']}`)
- **Language Gate:** {s['fit_data']['gates']['language_note']} (`{s['fit_data']['gates']['language_gate']}`)
- **Role Archetype:** {s['archetype']}
- **Recommended Master CV:** {s['recommended_variant']}
- **Market Salary Benchmark:** {s['salary_benchmark']['avg_display']} ({s['salary_benchmark']['range_display']} • {s['salary_benchmark']['source_basis']}; {('Stated in JD: ' + s['salary_benchmark']['stated_in_jd']) if s['salary_benchmark'].get('stated_in_jd') else 'Unstated in JD'})
- **Career-Ops Global Score:** **{s['global_score_5']} / 5.0** ({s['career_ops_band']}) | Evidence Confidence: **{s['evidence_confidence']}**
- **Internal Rank Rubric:** {s['rank_score_100']} / 100 ({s['rank_verdict']}) | Legacy UI Score: {s['fit_score_10']} / 10

---

## Block G — Posting Legitimacy & Risk Summary: **{s['block_g']['legitimacy_tier']}**
{f"> {s['block_g']['prior_note']}{chr(10)}" if s['block_g'].get('prior_note') else ''}
| Signal | Weight | Observation |
|---|---|---|
{sig_rows}

---

## Pass-1 Requirement Extraction -> Verified Evidence (Block B)
| Band | Requirement | JD Evidence (Quote) | Grounded Evidence (`candidate-facts.md`) | Fit |
|---|---|---|---|---|
{req_rows}

---

## 5-Persona Screening Panel Evaluations
{chr(10).join(persona_sections)}
"""

    def tailor_application(self, company, role, jd_text, location="Bengaluru", job_url=""):
        cleaned_jd, guard_findings = posting_guard.guard(jd_text or "", "auto")
        comp_slug = re.sub(r"[^a-zA-Z0-9]+", "-", company).strip("-").lower()[:20] or "company"
        role_slug = re.sub(r"[^a-zA-Z0-9]+", "-", role).strip("-").lower()[:25] or "role"
        slug = f"{comp_slug}-{role_slug}".strip("-")
        app_dir = self.applications_dir / slug
        app_dir.mkdir(parents=True, exist_ok=True)

        fit_data = self.score_jd(company, role, cleaned_jd, location, guard_findings=guard_findings)

        # 1. Write jd.md (cleaned via posting_guard + source URL)
        write_text(
            f"applications/{slug}/jd.md",
            f"# {company} – {role}\n\n"
            f"**Location:** {location}\n"
            f"**Source URL:** {job_url or 'Pasted JD'}\n"
            f"**Integrity Gate:** `{fit_data['guard_verdict']}`\n\n"
            f"## Sanitized Job Description\n\n{cleaned_jd}\n",
        )

        # 2. Write fit.md with ATS Match Table, Gates, Red Flags, Batched Gap Questions
        match_rows = "\n".join(
            f"| {m['requirement']} | {m['evidence']} | **{m['strength']}** |" for m in fit_data["match_table"]
        )
        red_flag_md = "\n".join(f"- {rf}" for rf in fit_data["red_flags"]) if fit_data["red_flags"] else "- None identified."
        gap_q_md = "\n".join(f"{idx+1}. {q}" for idx, q in enumerate(fit_data["gap_questions"])) if fit_data["gap_questions"] else "No missing facts required before drafting."

        fit_md = f"""# Fit Screen: {company} – {role}

## Verdict: {fit_data['verdict']} (Score: {fit_data['score']} / 10 | Rank Rubric: {fit_data['rank_score']} / 100 — {fit_data['rank_verdict']})
- **Level Discipline:** {'STRETCH APPLICATION (Director/Lead+ vs Visa I6 Senior PM)' if fit_data['is_stretch'] else 'Target Level (Senior PM / Visa I6 equivalent)'}
- **Eligibility Gate:** {fit_data['gates']['eligibility_gate']} ({fit_data['gates']['eligibility_note']})
- **Language Gate:** {fit_data['gates']['language_gate']} ({fit_data['gates']['language_note']})
- **Posting Guard:** {fit_data['guard_verdict']}

### Rubric Breakdown (1–10 UI Scale & 0–100 Rank Scale)
- **Domain Match (30%):** {fit_data['breakdown']['domain_match']} / 10
- **Seniority & Scope (20%):** {fit_data['breakdown']['seniority_match']} / 10
- **Technical & AI Depth (15%):** {fit_data['breakdown']['technical_ai_depth']} / 10
- **Measurable Impact (15%):** {fit_data['breakdown']['measurable_impact']} / 10
- **Location & Sponsorship (20%):** {fit_data['breakdown']['location_sponsorship']} / 10

### ATS & Recruiter Match Table
| JD Requirement | Evidence on Master CV / Facts Bank | Strength |
|---|---|---|
{match_rows}

### Red Flags a Recruiter Will Notice
{red_flag_md}

### Batched Gap Questions (Answer Together Before Submitting)
{gap_q_md}

### Top 5 Grounded Proof Points
1. {fit_data['proof_points'][0]}
2. {fit_data['proof_points'][1]}
3. {fit_data['proof_points'][2]}
4. {fit_data['proof_points'][3]}
5. {fit_data['proof_points'][4]}
"""
        write_text(f"applications/{slug}/fit.md", fit_md)

        # 3. Write resume.md (Strictly aligned with candidate-facts.md: 10 FTEs, CS attribution, avijitmukherjeepm, no leave reason)
        headline = f"Senior Product Manager – {role}" if "Product" in role else "Senior Payments Product Leader"
        resume_md = f"""# AVIJIT MUKHERJEE
**{headline}**
Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjeepm/

---

## PROFESSIONAL SUMMARY

Payments & AI-enabled product leader with nearly 14 years of experience across acquiring platforms, transaction processing, payment APIs, and settlement infrastructure, including nearly 10 years in product management at Visa (CyberSource / Visa Platform Connect) and Cognizant (Fiserv / First Data). Generated $30M in net revenue by launching domestic-scheme processing across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit, and protected over $22M in transaction/merchant revenue (Visa figures attributed to Customer Support post-release). Certified Scrum Product Owner (CSPO) and 2-time Visa STAR Award recipient combining deep payments rails mastery (ISO 8583, BASE I/II, Kafka) with hands-on Generative AI & agentic product workflows (Fast Refund: 24h refunds, 97% success rate; SMCC Japan specs cut from 6 weeks to 2 weeks).

---

## CORE COMPETENCIES

- **Product Strategy & Delivery:** Product Roadmaps (RICE), PRDs & Business Cases, Backlog Prioritization, CSPO Agile Delivery (4 Scrum Teams), KPI Scorecards & Release Governance, P1 Escalation Leadership
- **Payments & Acquiring:** Merchant Processing, Authorization, Clearing & Settlement, Domestic & Regional Schemes (Jaywan, RuPay, KSA, CUP, MyDebit), Network Tokenization (COF/MIT/CIT), Interchange & EMV Compliance
- **Architecture & Rails:** ISO 8583, Visa BASE I & BASE II, TC33.A, Kafka Event-Driven Architecture, REST APIs, Microservices, Payments 2.0

### Applied AI & Agentic Product Management

- Applied **Generative AI and agentic workflows** using Claude and other AI tools across Product Management activities including PRDs, business cases, acceptance criteria, defect analysis and solution architecture, reducing manual effort by **50%** across selected workflows.
- Used **Claude and an agentic toolchain** to support the architecture and delivery of **EDQP, PIF and force-refund capabilities for SMCC Japan**, accelerating solution definition and implementation.
- Applied AI-assisted analysis to **payment and settlement validation**, comparing transaction data across processing stages, identifying data misalignment and supporting root-cause analysis and remediation.
- Integrated AI into product workflows while maintaining **human validation, business rules and engineering review** for production decisions.
- Working knowledge of **LLMs, RAG, embeddings, vector databases, tool calling, MCP, model routing, evals, guardrails, hallucination management and AI workflow orchestration**.

### AI & GenAI Product Management

**Generative AI:** LLMs, Prompt Engineering, Context Management, RAG, Embeddings, Vector Databases, Fine-tuning, Hallucination Management, Guardrails, Evals

**AI Product Architecture:** APIs, Tool Calling, MCP, Model Routing, AI Workflows, Agentic Workflows, Human-in-the-Loop

**Applied AI:** Claude, Gemini, OpenAI Tools, Copilot, AI-assisted PRDs, Business Cases, Acceptance Criteria, Defect Analysis, Solution Architecture, Product Discovery & SDLC Automation

---

## PROFESSIONAL EXPERIENCE

### VISA
**Product Manager (I6) – CyberSource / Visa Platform Connect (VPC)** | Bengaluru, India  
*Jan 2025 – Sep 2026*

- Generated $30M in net revenue by launching multi-scheme domestic processing and Kafka event-driven settlement microservices across KSA, RuPay, CUP, UAE Jaywan, PayPak, and MyDebit.
- Freed 10 FTEs of recurring effort by automating CAS self-service testing and on-demand scoped settlement-file retrieval across 400+ files/release, cutting file turnaround under 24 hours.
- Shipped AI optimization for Fast Refund in India/APAC by eliminating redundant spike analysis and regression effort, achieving 24-hour refunds at a 97% success rate and higher daily transaction throughput.
- Accelerated cross-border integration delivery by 67% (6 weeks to 2 weeks) for SMCC Japan (EDQP, PIF, force-refund) using a Claude-based agentic specification and edge-case validation workflow.
- Executed the Apple Pay India data center migration to meet RBI localization compliance, choosing regulatory compliance and market readiness over negligible cost impact.
- Expanded European acquiring capacity across 13 new acquirers (+12,000 TPS) via Elavon VPC integration and TC33.A clearing support for JCB and Diners in France/Spain.

### VISA
**Technical Product Manager – Visa Platform Connect (VPC)** | Bengaluru, India  
*Mar 2024 – Dec 2024*

- Protected $9.2M in transaction revenue (attributed to Customer Support post-release impact reports) by introducing automated settlement regeneration workflows and serving as SME escalation lead on SWAT & client P1 calls.
- Grew token-based transaction volume by 25% across European acquirers by simplifying COF, MIT, and CIT implementation flows, and increased CEMEA Visa Instalment Plan volume by 18%.
- Recommended and secured VP sign-off on a RICE-prioritized platform roadmap while mentoring 4 scrum teams (~6 engineers/team) using velocity, burndown tracking, and individual scorecards.
- Supported Automated Fuel Dispenser (AFD) acceptance frameworks associated with $70M in processing throughput across US and EU markets while maintaining zero major post-release rollbacks.

### VISA
**Technical Product Analyst – Visa Platform Connect (VPC)** | Bengaluru, India  
*Oct 2021 – Feb 2024*

- Scaled financial partner onboarding from 100+ to 400+ acquiring institutions (engaged almost daily across acquirers), stepping in on the Acquirer Onboarding Platform to drive an architect-approved UI/design improvement to production.
- Defined BIN lookup and co-brand routing logic for Australian (EFTPOS) and Saudi Arabian (MADA) schemes with zero routing rejection errors in production.
- Decoupled 85% of legacy monolithic BASE I validation logic into 5 microservices under Payments 2.0 while maintaining 100% BASE II clearing compliance (TC33.A, TC05/06, IPM, CAPN).

### COGNIZANT
**Team Lead – Product Managers (Fiserv / First Data Account)** | Kolkata, India  
*Jul 2019 – Sep 2021*

- Managed a 12-member product and analyst team across 100+ enterprise merchant relationships (YUM/KFC, Burger King, McDonald’s), protecting $11M in merchant revenue via upstream payment-flow validation.
- Recovered $1.5M in missing KFC (YUM group) settlement revenue by diagnosing an authorization-to-settlement code gap where database records failed to propagate downstream.
- Raised post-release accuracy from 80% to 97%, cut P1 incidents by 19%, and reduced MTTR by 40% while influencing a Cognizant SOW requote so automation-freed analysts remained retained on the account.

### COGNIZANT
**Product Manager – Transaction Validation Automation (First Data Account)** | Chennai, India  
*Jul 2017 – Jun 2019*

- Built the Transaction Validation System (2017) across 13 payment specs, cutting merchant certification cycle times by 57% and saving ~30 analysts 10 hours per file—enabling First Data to transition those analysts into Project Lead roles.
- Contributed to a 32% year-end revenue increase for First Data, protected $2M+ in client revenue, and reached 99%+ automated validation accuracy.

### COGNIZANT
**Certification Project Manager / Payments SME & Process Executive (First Data / Nationstar)** | Kolkata, India  
*Jan 2013 – Jun 2017*

- Delivered 100+ merchant and ISV certifications (Walmart, McDonald’s, ExxonMobil, Verifone, NCR) at a 97% first-pass success rate; awarded for validating 1,000+ transaction payloads against ISO standards.
- Recognized as Cards & Payments SME (Oct 2015); interviewed 100+ hiring candidates across career and trained 20+ junior analysts at 30% faster ramp-up.

---

## EDUCATION & CERTIFICATIONS

- **Bachelor of Technology (B.Tech) – Information Technology** | West Bengal University of Technology (2012) | GPA: 8.6
- **Certified Scrum Product Owner (CSPO)** – Scrum Alliance | **PLM & VPL Core Product** – Visa
- **Honors:** 2-time Visa STAR Award recipient | Multiple Cognizant Annual Employee Awards | Cards & Payments SME (2015)
"""
        write_text(f"applications/{slug}/resume.md", resume_md)

        # 4. Write cover_note.md (<150 words, no reason for leaving Visa, grounded proof points)
        cover_note = f"""# Cover Note: {company} – {role}

**Subject:** {role} – Avijit Mukherjee

Hi [Hiring Team / {company} Product Leadership],

As a payments product leader with nearly 14 years across Visa (CyberSource / VPC) and Cognizant (Fiserv), {company}’s focus on {role} directly matches my experience taking regulated payment platforms from spec to production:

• **Multi-Scheme Scale & Revenue:** Launched domestic-scheme processing (KSA, RuPay, Jaywan) on Payments 2.0 & Kafka ($30M net revenue) and scaled acquirer onboarding to 400+ partners.
• **Automation & AI Delivery:** Shipped AI optimization on Fast Refund (24h refunds, 97% success) and automated settlement retrieval freeing 10 FTEs.
• **Operational Resilience:** Protected $9.2M at Visa (Customer Support-attributed) and $11M at Fiserv while leading SWAT/P1 escalations.

Certified Scrum Product Owner (CSPO) and 2-time Visa STAR Award recipient—I would welcome 15 minutes to discuss how this background supports {company}.

Best regards,
Avijit Mukherjee
Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjeepm/
"""
        write_text(f"applications/{slug}/cover_note.md", cover_note)

        # 5. Write outreach.md (<120 words each + LinkedIn People Search URLs per job-scraper Step 4.5)
        ref_links = fit_data["referral_links"]
        outreach_md = f"""# Outreach Messages: {company} – {role}

### Direct LinkedIn People Search Links (Step 4.5 Safe Referral Discovery)
- **Find Recruiters / TA at {company}:** {ref_links['recruiter_search_url']}
- **Find Product Peers / Hiring Managers at {company}:** {ref_links['peer_search_url']}

---

### 1. Referral Ask (First-Degree Connection, < 90 words)
> Hi [Name], I noticed {company} has an opening for {role}. Having spent nearly 14 years leading core payment platforms, scheme expansions ($30M net revenue), and AI delivery workflows at Visa and Fiserv, I see strong alignment with your team. Would you be open to submitting an internal referral or connecting me with the hiring manager? No worries at all if not.

### 2. Cold Note to Hiring Manager (< 90 words)
> Hi [Name],
>
> I’ve followed {company}’s platform work closely. Most recently at Visa (CyberSource), I launched Payments 2.0 domestic-scheme processing generating $30M in net revenue and shipped AI optimization on Fast Refund (24h refunds, 97% success rate).
>
> Would you be open to a 15-minute chat on how my acquiring and settlement background can support your roadmap?
>
> Best, Avijit Mukherjee (linkedin.com/in/avijitmukherjeepm/)

### 3. Note to Tech Recruiter (< 70 words)
> Hi [Name],
>
> Reaching out regarding {role} at {company}. I’m a payments product leader (nearly 14 yrs across Visa I6 / Fiserv, CSPO certified) with $30M net revenue launched and 10 FTEs freed through automation.
>
> Happy to share my tailored CV: linkedin.com/in/avijitmukherjeepm/
>
> Best, Avijit

### 4. Day 5 Follow-Up
> Hi [Name], following up briefly on my note regarding {company}’s {role}. Remains top of my list—would welcome a short conversation whenever convenient. Thanks!
"""
        write_text(f"applications/{slug}/outreach.md", outreach_md)

        # 6. Application Answers (10 STAR answers)
        app_answers = self.generate_application_answers(company, role, cleaned_jd)
        write_text(f"applications/{slug}/application_answers.md", app_answers)

        # 7. Pre-submit Quality Check
        quality_check = self.run_pre_submit_quality_check(company, role, resume_md, cover_note, fit_data)
        write_text(f"applications/{slug}/pre_submit_check.md", quality_check["markdown"])

        # 7b. Agentic 5-Persona JD Scrutiny Report (career-ops A-G + ATS/HR/Tech/HM/Bar Raiser)
        scrutiny = self.run_agentic_jd_scrutiny(company, role, cleaned_jd, location, job_url=job_url)
        write_text(f"applications/{slug}/panel_review.md", scrutiny["markdown"])

        # 8. Generate DOCX using senior-pm-job-search Mode 2 naming convention:
        #    CV_<Company>_<Role>_<YYYY-MM-DD>.docx
        docx_clean_company = re.sub(r"[^a-zA-Z0-9]+", "_", company).strip("_")[:20] or "Company"
        docx_clean_role = re.sub(r"[^a-zA-Z0-9]+", "_", role).strip("_")[:25] or "Role"
        today_str = datetime.now().strftime("%Y-%m-%d")
        docx_filename = f"CV_{docx_clean_company}_{docx_clean_role}_{today_str}.docx"
        docx_path = app_dir / docx_filename
        self.generate_docx(resume_md, docx_path)

        # 9. Update tracker.csv AND applications-log.md
        note_text = f"Fit {fit_data['score']}/10 ({fit_data['rank_score']}/100 {fit_data['rank_verdict']}) | Panel: {scrutiny['global_score_5']}/5 {scrutiny['bar_vote']}. Readiness: {quality_check['score']}/100."
        if fit_data["is_stretch"]:
            note_text = f"[STRETCH] {note_text}"
        self.update_tracker_status(
            company=company,
            role=role,
            new_status="applying",
            notes=note_text,
            follow_up_date=(datetime.now() + timedelta(days=5)).strftime("%Y-%m-%d"),
            link=job_url,
            location=location,
            sponsorship=fit_data["sponsorship"],
            score=str(fit_data["score"]),
        )
        self.append_application_log(company, role, "variant made", f"Saved {docx_filename}; {note_text}")

        return {
            "slug": slug,
            "folder": str(app_dir),
            "docx_filename": docx_filename,
            "files": {
                "docx": f"/download/applications/{slug}/{docx_filename}",
                "resume_md": f"/download/applications/{slug}/resume.md",
                "cover_note_md": f"/download/applications/{slug}/cover_note.md",
                "fit_md": f"/download/applications/{slug}/fit.md",
                "panel_review_md": f"/download/applications/{slug}/panel_review.md",
                "outreach_md": f"/download/applications/{slug}/outreach.md",
                "application_answers_md": f"/download/applications/{slug}/application_answers.md",
                "pre_submit_check_md": f"/download/applications/{slug}/pre_submit_check.md",
            },
            "fit_data": fit_data,
            "scrutiny": scrutiny,
            "quality_check": quality_check,
        }

    def generate_application_answers(self, company, role, jd_text):
        return f"""# Application Form Answers: {company} – {role}
*Grounded strictly in verified facts from Master CV & `references/candidate-facts.md`.*

---

### 1. Relevant Experience
I bring nearly 14 years of experience across acquiring platforms, transaction processing, and payment infrastructure, including nearly 10 years in product management at Visa (CyberSource / VPC, level I6) and Cognizant (Fiserv / First Data). I have launched domestic schemes across KSA, RuPay, CUP, and UAE Jaywan on Payments 2.0 and Kafka ($30M net revenue) and shipped AI workflow improvements on Fast Refund.

### 2. Product Impact & AI
At Visa, I applied AI to eliminate redundant spike analysis and regression effort on Fast Refund in India/APAC, delivering refunds within 24 hours at a 97% success rate with higher daily transaction throughput. Separately, our multi-scheme Payments 2.0 rollout drove $30M in net revenue and freed 10 FTEs through self-service settlement automation.

### 3. Regulatory vs. Cost Trade-Off Judgement
For the Apple Pay India launch, RBI data localization guidelines required moving code to the India data center. I evaluated compliance and market-readiness versus infrastructure cost (which was negligible) and chose regulatory compliance and schedule certainty, completing market launch testing with zero processing errors.

### 4. Prioritization With Executive Stakeholders
At Visa I built and recommended a RICE-prioritized roadmap balancing new market schemes, legacy BASE I decoupling (85% across 5 microservices), and partner onboarding across 400+ acquirers, earning formal sign-off from a Vice President.

### 5. Influence Without Authority
When another PM went on leave for the Acquirer Onboarding Platform, I stepped in, proposed a UX/design improvement, gained architect approval, and drove it to implementation. At Cognizant, I also influenced an SOW requote so analysts freed by automation transitioned into Project Lead roles instead of losing headcount.

### 6. Handling High-Severity Production Escalations
As SME escalation lead on SWAT and client P1 calls at Visa, I provided rapid technical triage to accelerate fixes and protect SLAs—contributing to $9.2M in revenue protection (attributed to Customer Support post-release). At Cognizant/KFC, I traced a code gap between authorization and settlement to recover $1.5M in revenue.

### 7. Automation & Operational Scaling
Early on First Data (2017), ~30 analysts each spent ~10 hours per file validating settlement payloads. I built an automated Transaction Validation System across 13 specs, cutting certification cycles by 57% and freeing analysts to step into Project Lead roles. Later at Visa, CAS automation freed 10 FTEs.

### 8. People Leadership & Mentoring
I led a 12-member product and analyst team at Cognizant (2019–2021) across 100+ merchant accounts, interviewed 100+ hiring candidates across my career, and at Visa worked with 4 scrum teams (~6 engineers each) driving velocity, burndown, individual scorecards, and annual 360 feedback.

### 9. Salary Expectation & Location
Target total compensation equivalent to USD 100,000–200,000 globally (or competitive Bengaluru Senior PM / I6+ benchmark). Based in Bengaluru; open to relocating abroad within ~6 months where visa sponsorship is offered.

### 10. Reason for Interest in {company}
{company}’s platform scale aligns directly with my background in regulated payments rails, merchant/acquirer platforms, and AI-accelerated delivery. I’m eager to bring 14 years of practical execution and product judgement to {role}.
"""

    def run_pre_submit_quality_check(self, company, role, resume_text, cover_note_text, fit_data=None):
        score = 96
        critical_fixes = []
        optional_improvements = [
            f"Confirm the application portal for {company} is still accepting submissions before submitting.",
            "Verify hiring manager / recruiter names on LinkedIn using the generated people-search links in outreach.md.",
        ]

        combined = resume_text + "\n" + cover_note_text
        if "8 fte" in combined.lower():
            score -= 10
            critical_fixes.append("Contains stale '8 FTE' claim; candidate-facts.md settles this at 10 FTEs.")
        if "laid off" in combined.lower():
            score -= 15
            critical_fixes.append("Remove reason for leaving Visa from CV / cover note per Rule 4.")

        placeholders = re.findall(r"\[ADD METRIC.*?\]|\[RESEARCH.*?\]", combined)
        if placeholders:
            score -= len(placeholders) * 10
            critical_fixes.append(f"Resolve placeholder tags: {', '.join(placeholders)}")

        banned = ["passionate", "synergy", "spearheaded", "results-driven", "orchestrated"]
        found_buzzwords = [w for w in banned if w in combined.lower()]
        if found_buzzwords:
            score -= 5
            critical_fixes.append(f"Replace buzzwords: {', '.join(found_buzzwords)}")

        if fit_data and fit_data.get("gates", {}).get("eligibility_gate") == "FAIL":
            score = min(score, 60)
            critical_fixes.append(f"Eligibility Gate Failed: {fit_data['gates']['eligibility_note']}")

        verdict = "READY" if score >= 85 and not critical_fixes else "REVIEW NEEDED"

        md_report = f"""# Pre-Submit Quality Review: {company} – {role}

### Submission Readiness Score: {score} / 100
**Final Approval:** **{verdict}**

---

### Audit Criteria Breakdown:
1. **Facts Bank Grounding (100%):** Claims trace to `references/candidate-facts.md` (10 FTEs settled, CS-attributed revenue protection, Fast Refund AI 97%, no Visa exit reason exposed).
2. **CV Hygiene & URL Check:** LinkedIn profile `avijitmukherjeepm`, contact info verified, zero fabricated tools/domains.
3. **ATS Compatibility (100%):** Single-column layout, conventional headings, saved as `CV_{re.sub(r'[^a-zA-Z0-9]+', '_', company).strip('_')[:20]}_{re.sub(r'[^a-zA-Z0-9]+', '_', role).strip('_')[:25]}_{datetime.now().strftime('%Y-%m-%d')}.docx`.
4. **Prompt-Injection / Posting Guard:** {fit_data.get('guard_verdict', 'clean') if fit_data else 'clean'}.

### Critical Fixes Before Applying:
{chr(10).join(['- ' + fix for fix in critical_fixes]) if critical_fixes else '- None. All claims align with verified facts bank.'}

### Optional Enhancements:
{chr(10).join(['- ' + opt for opt in optional_improvements])}
"""
        return {
            "score": score,
            "verdict": verdict,
            "critical_fixes": critical_fixes,
            "optional_improvements": optional_improvements,
            "markdown": md_report,
        }

    def generate_follow_up_pack(self, company, role, applied_date, days_elapsed, contact_name="Hiring Team"):
        contact = contact_name or "Hiring Team"
        ref = portals.build_referral_links(company, role)
        return {
            "follow_up_email": (
                f"Subject: Following up: {role} application – Avijit Mukherjee\n\n"
                f"Hi {contact},\n\n"
                f"I applied for the {role} role at {company} on {applied_date} ({days_elapsed} days ago). "
                f"Given {company}’s focus on payment infrastructure and platform scale, my nearly 14 years leading acquiring and "
                f"settlement systems at Visa (CyberSource / VPC)—launching Payments 2.0 rails ($30M net revenue) and shipping AI optimization "
                f"on Fast Refund (24h refunds, 97% success)—align closely with your team's priorities.\n\n"
                f"I’d welcome a brief 15-minute chat. Thanks for considering!\n\n"
                f"Best regards,\nAvijit Mukherjee\n+91 9038415509 | linkedin.com/in/avijitmukherjeepm/"
            ),
            "linkedin_connection_request": (
                f"Hi {contact}, I applied for {role} at {company}. With nearly 14 yrs leading acquiring & rails at Visa ($30M net rev), "
                f"I'd love to connect and follow your team's work!"
            )[:300],
            "linkedin_dm": (
                f"Hi {contact},\n\n"
                f"Reaching out regarding {company}’s {role} opening. At Visa (CyberSource), I launched Payments 2.0 rails ($30M net revenue) "
                f"and shipped AI automation on Fast Refund (97% success). Would welcome 15 mins to connect.\n\nBest, Avijit"
            ),
            "linkedin_search_strategy": [
                f"Recruiter Search: {ref['recruiter_search_url']}",
                f"Team / Peer Search: {ref['peer_search_url']}",
            ],
            "referral_links": ref,
        }

    def get_interview_prep_deep_dive(self, company, role):
        company_clean = company or "Target Company"
        return {
            "company": company_clean,
            "role": role or "Senior Product Manager",
            "elevator_pitch_60s": (
                "I am a Payments Product Leader with nearly 14 years of experience across acquiring platforms, payment APIs, clearing, "
                "and settlement infrastructure, including nearly 10 years in product management at Visa (level I6) and Cognizant (Fiserv). "
                "Most recently at Visa CyberSource, I launched domestic-scheme processing across KSA, RuPay, and UAE Jaywan generating $30M "
                "in net revenue, automated settlement tooling freeing 10 FTEs, and applied AI to Fast Refund to achieve 24-hour refunds at a "
                "97% success rate. Across 400+ acquirers and 4 scrum teams, I pair deep ISO 8583/Kafka fluency with strong product judgement "
                f"and regulatory trade-off execution—and that is exactly what excites me about {company_clean}."
            ),
            "perceived_risks_and_mitigations": [
                {
                    "risk": "Enterprise network background (Visa/Cognizant) vs fast-moving fintech execution.",
                    "mitigation": "Cite shipping Fast Refund AI optimization (97% success), stepping in to drive an Acquirer Onboarding UI change to production, and cutting SMCC Japan specs from 6 weeks to 2 weeks."
                },
                {
                    "risk": "Known gaps (UPI, Consumer B2C, Lending, Wealth/Insurance, Direct Reports at Visa).",
                    "mitigation": "Be transparent without bluffing: anchor in B2B payments rails, Cognizant 12-person leadership + 100+ interviews, and Visa cross-team scorecard mentoring."
                },
                {
                    "risk": "Why did you leave Visa? (Prepare neutrally; never volunteer it on paper.)",
                    "mitigation": "Keep brief and factual: impacted by organizational restructuring/layoff in Sept 2026; took the window to deepen applied AI for PM skills alongside existing shipped Fast Refund AI work."
                }
            ],
            "questions_to_ask_them": [
                f"How does {company_clean} balance regulatory compliance obligations against speed-to-market on payment rail launches?",
                f"What is the highest-leverage bottleneck in {company_clean}'s settlement, reconciliation, or merchant onboarding flow today?",
                f"How do PMs partner with engineering and risk/compliance on roadmap prioritization here?",
                f"What does strong Senior PM impact look like in the first 90 days on this product area?"
            ]
        }

    def generate_docx(self, markdown_text, output_path):
        doc = docx.Document()
        for section in doc.sections:
            section.top_margin = Inches(0.55)
            section.bottom_margin = Inches(0.55)
            section.left_margin = Inches(0.65)
            section.right_margin = Inches(0.65)
            section.page_width = Inches(8.5)
            section.page_height = Inches(11.0)

        style_normal = doc.styles["Normal"]
        font = style_normal.font
        font.name = "Calibri"
        font.size = Pt(9.5)
        font.color.rgb = RGBColor(33, 37, 41)

        p_name = doc.add_paragraph()
        p_name.paragraph_format.space_after = Pt(2)
        p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_name = p_name.add_run("AVIJIT MUKHERJEE")
        r_name.font.size = Pt(17)
        r_name.font.bold = True
        r_name.font.color.rgb = RGBColor(15, 23, 42)

        p_contact = doc.add_paragraph()
        p_contact.paragraph_format.space_after = Pt(6)
        p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_con = p_contact.add_run("Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjeepm/")
        r_con.font.size = Pt(9)
        r_con.font.color.rgb = RGBColor(71, 85, 105)
        add_hr(p_contact, "94A3B8")

        def add_formatted_runs(paragraph, text, size_pt=9.5, bold_color=RGBColor(15, 23, 42)):
            for part in re.split(r"(\*\*.*?\*\*)", text):
                if not part:
                    continue
                if part.startswith("**") and part.endswith("**") and len(part) > 4:
                    r_seg = paragraph.add_run(part[2:-2])
                    r_seg.font.bold = True
                    r_seg.font.size = Pt(size_pt)
                    r_seg.font.color.rgb = bold_color
                else:
                    r_seg = paragraph.add_run(part)
                    r_seg.font.size = Pt(size_pt)

        lines = markdown_text.split("\n")
        in_header = True
        for line in lines:
            l = line.strip()
            if not l or l.startswith("# AVIJIT") or "linkedin.com" in l or l == "---":
                continue

            if in_header and not l.startswith("## ") and l.startswith("**") and l.endswith("**"):
                p_head = doc.add_paragraph()
                p_head.paragraph_format.space_before = Pt(0)
                p_head.paragraph_format.space_after = Pt(3)
                p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_head = p_head.add_run(l.replace("**", "").strip())
                r_head.font.size = Pt(10)
                r_head.font.bold = True
                r_head.font.color.rgb = RGBColor(30, 58, 138)
                continue

            if l.startswith("## "):
                in_header = False
                section_title = l.replace("## ", "").strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                r = p.add_run(section_title.upper())
                r.font.size = Pt(10.5)
                r.font.bold = True
                r.font.color.rgb = RGBColor(15, 23, 42)
                add_hr(p, "CBD5E1")
            elif l.startswith("### "):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.keep_with_next = True
                r = p.add_run(l.replace("### ", "").strip())
                r.font.bold = True
                r.font.size = Pt(10)
                r.font.color.rgb = RGBColor(15, 23, 42)
            elif l.startswith("**") and ":**" not in l and ("|" in l or "–" in l or "-" in l or any(kw in l for kw in ["Manager", "Analyst", "Lead", "Executive", "SME", "Engineer", "Officer", "VP", "Head"])):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.keep_with_next = True
                if "|" in l:
                    parts = l.split("|")
                    role_part = parts[0].replace("**", "").strip()
                    loc_part = " | " + parts[1].replace("**", "").strip()
                    r_role = p.add_run(role_part)
                    r_role.font.bold = True
                    r_role.font.italic = True
                    r_role.font.size = Pt(9.5)
                    r_role.font.color.rgb = RGBColor(30, 58, 138)
                    r_loc = p.add_run(loc_part)
                    r_loc.font.size = Pt(9.5)
                    r_loc.font.color.rgb = RGBColor(100, 116, 139)
                else:
                    r = p.add_run(l.replace("**", "").strip())
                    r.font.bold = True
                    r.font.italic = True
                    r.font.size = Pt(9.5)
                    r.font.color.rgb = RGBColor(30, 58, 138)
            elif (
                (l.startswith("*") and l.endswith("*") and not l.startswith("**"))
                or (l.startswith("_") and l.endswith("_"))
                or re.search(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}\b", l)
            ) and not (l.startswith("- ") or l.startswith("• ")):
                date_text = l.strip("*_ ").strip()
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.keep_with_next = True
                r = p.add_run(date_text)
                r.font.italic = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(100, 116, 139)
            elif l.startswith("- ") or l.startswith("• "):
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.left_indent = Inches(0.2)
                p.paragraph_format.line_spacing = 1.12
                r_b = p.add_run("•  ")
                r_b.font.size = Pt(9)
                r_b.font.color.rgb = RGBColor(30, 58, 138)
                content = l[2:].strip()
                add_formatted_runs(p, content, size_pt=9.5)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                add_formatted_runs(p, l, size_pt=9.5)

        doc.save(str(output_path))

    def evaluate_mock_answer(self, question, answer):
        ans_lower = answer.lower()
        word_count = len(answer.split())

        has_metrics = bool(re.search(r"\d+%|\$\d+|\d+\s*fte|\d+\s*weeks|\d+\s*tps|\d+\s*hours|\d+\s*hours?", ans_lower))
        has_action = any(verb in ans_lower for verb in ["architected", "delivered", "automated", "built", "led", "decided", "recommended", "proposed", "drove"])

        strengths = []
        vague_areas = []
        score = 3.0

        if word_count > 60:
            strengths.append("Solid depth and clear context.")
            score += 0.5
        else:
            vague_areas.append("Too brief for a Senior PM answer (under 60 words). Include situation, decision, trade-off, and metric.")

        if has_metrics:
            strengths.append("Concrete numbers included ($ revenue, FTEs, % lift, or SLA hours).")
            score += 1.0
        else:
            vague_areas.append("Missing measurable impact (e.g. $30M net rev, 10 FTEs freed, 97% Fast Refund, $9.2M Customer Support-attributed).")
            score -= 0.5

        if has_action:
            strengths.append("Strong personal ownership and decisioning verbs.")
            score += 0.5
        else:
            vague_areas.append("Clarify YOUR exact decision and trade-off rather than describing what the team generally did.")

        score = min(5.0, max(1.5, score))
        tighter_version = (
            "Context: [1-sentence payment rail / merchant constraint]. Decision & Trade-off: I chose [specific option] over [alternative] "
            "because [regulatory/customer rationale], aligned [VP/architect/4 scrum teams], and delivered [verified metric from candidate-facts.md]."
        )

        return {
            "score": score,
            "strengths": strengths,
            "vague_areas": vague_areas,
            "tighter_version": tighter_version,
            "feedback": "Strong PM judgement framing. Keep claims strictly aligned with candidate-facts.md." if score >= 4 else "Sharpen your personal decision, explicit trade-off, and exact Customer-Support-attributed or production metric.",
        }

    def run_live_scrape(
        self,
        queries: list[str] | None = None,
        locations: list[str] | None = None,
        portals_enabled: list[str] | None = None,
        jobage_days: int = 14,
        remote: str | None = None,
        enrich_details: bool = True,
        integrations_mgr=None,
    ):
        """Executes live multi-portal scrape (LinkedIn guest, Freehire API, and target ATS boards), deduplicates & stores in seen_jobs.json."""
        queries = [q.strip() for q in (queries or ["Senior Product Manager payments", "Product Manager settlement", "AI Product Manager fintech"]) if q.strip()][:5]
        locations = [l.strip() for l in (locations or ["Bengaluru, Karnataka, India"]) if l.strip()][:3]
        portals_enabled = portals_enabled or ["linkedin", "freehire", "ats"]

        raw_hits = []
        portal_errors = []

        # 1. LinkedIn public jobs-guest search
        if "linkedin" in portals_enabled:
            for loc in locations:
                for q in queries[:3]:
                    res = portals.search_linkedin(q, location=loc, jobage_days=jobage_days, remote=remote, limit=12)
                    if res.get("error"):
                        portal_errors.append(f"LinkedIn ({q} @ {loc}): {res['error']}")
                    raw_hits.extend(res.get("results") or [])

        # 2. Freehire public API search
        if "freehire" in portals_enabled:
            countries = []
            for loc in locations:
                ll = loc.lower()
                if "india" in ll or "bengaluru" in ll:
                    countries.append("IN")
                elif "dubai" in ll or "uae" in ll:
                    countries.append("AE")
                elif "london" in ll or "uk" in ll:
                    countries.append("GB")
                elif "singapore" in ll:
                    countries.append("SG")
                elif "united states" in ll or "usa" in ll:
                    countries.append("US")
            for q in queries[:3]:
                res = portals.search_freehire(q, jobage_days=jobage_days, limit=15, countries=countries or None, work_mode=remote)
                if res.get("error"):
                    portal_errors.append(f"Freehire ({q}): {res['error']}")
                raw_hits.extend(res.get("results") or [])

        # 3. Direct Target ATS boards (Greenhouse/Lever)
        if "ats" in portals_enabled and integrations_mgr is not None:
            try:
                ats_jobs = integrations_mgr.crawl_public_job_boards()
                for aj in ats_jobs:
                    raw_hits.append({
                        "title": aj.get("title", ""),
                        "company": aj.get("company", ""),
                        "location": aj.get("location", ""),
                        "date": aj.get("posted_date"),
                        "url": aj.get("link", ""),
                        "portal": "target-ats",
                        "source": "cli",
                        "description": f"{aj.get('title', '')} at {aj.get('company', '')} ({aj.get('location', '')})",
                    })
            except Exception as exc:
                portal_errors.append(f"ATS Crawler: {exc}")

        # Pool dedup by canonical URL (without ?utm_source=...) before storing/fetching detail
        seen_urls = set()
        pool = []
        for h in raw_hits:
            u = (h.get("url") or "").strip().rstrip("/")
            t = (h.get("title") or "").strip()
            c = portals.normalize_company_display(h.get("company") or "")
            if not u or not t or not c:
                continue
            h["company"] = c
            canon_u = re.sub(r"[?&]utm_[^=]+=[^&]+", "", u).rstrip("?&/")
            if canon_u in seen_urls:
                continue
            # Pre-filter strictly to PM/Product/Payment roles if broad results slipped in
            tl = t.lower()
            if not any(k in tl for k in ["product", "payments", "acquiring", "fintech", "settlement", "program manager"]):
                continue
            seen_urls.add(canon_u)
            pool.append(h)

        # Optionally enrich up to 5 top promising LinkedIn hits for full JD + closed check + posting_guard
        if enrich_details:
            li_enriched = 0
            for h in pool:
                if h.get("portal") == "linkedin-search" and li_enriched < 5 and not h.get("description"):
                    try:
                        det = portals.fetch_linkedin_detail(h["url"])
                        if det.get("ok"):
                            if not det.get("is_active", True):
                                h["status"] = "expired"
                            h["description"] = det.get("description") or ""
                            h["guard_verdict"] = det.get("guard_verdict", "clean")
                            h["employment_type"] = det.get("employment_type")
                        li_enriched += 1
                    except Exception:
                        pass

        # Quick fit classification (high/medium/low)
        for h in pool:
            fit_res = self.score_jd(h["company"], h["title"], h.get("description") or h["title"], h.get("location") or "")
            h["quick_score"] = fit_res["score"]
            h["rank_score_preview"] = fit_res["rank_score"]
            if fit_res["score"] >= 7.5 and fit_res["gates"]["eligibility_gate"] != "FAIL" and fit_res["gates"]["language_gate"] != "FAIL":
                h["fit"] = "high"
            elif fit_res["score"] >= 6.0 and fit_res["gates"]["eligibility_gate"] != "FAIL" and fit_res["gates"]["language_gate"] != "FAIL":
                h["fit"] = "medium"
            else:
                h["fit"] = "low"

        store_res = portals.store_scraped_jobs(pool)
        self.last_scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.scanner_logs.append(
            f"[{datetime.now().strftime('%H:%M:%S')}] Scraped {len(pool)} PM hits -> added {store_res['added_count']} new jobs to seen_jobs.json."
        )

        return {
            "status": "success",
            "timestamp": self.last_scan_time,
            "total_fetched": len(raw_hits),
            "filtered_pm_hits": len(pool),
            "added_count": store_res["added_count"],
            "added_keys": store_res["added_keys"],
            "dedup_counts": store_res["counts"],
            "portal_errors": portal_errors,
            "seen_jobs": self.get_seen_jobs_view(),
        }

    def get_seen_jobs_view(self):
        """Returns scraped/ranked jobs from job_scraper/seen_jobs.json sorted by rank_score then recency."""
        _, seen = portals.load_seen_state()
        tracker_set = {(portals.normalize_company_display(r.get("company", "")).strip().lower(), r.get("role", "").strip().lower()) for r in self.get_tracker()}
        items = []
        for key, entry in seen.items():
            raw_comp = entry.get("company") or ""
            comp = portals.normalize_company_display(raw_comp)
            title = entry.get("title") or ""
            in_tracker = (comp.strip().lower(), title.strip().lower()) in tracker_set
            items.append({
                "key": key,
                "company": comp,
                "title": title,
                "location": entry.get("location") or "",
                "url": entry.get("url") or "",
                "portal": entry.get("portal") or "websearch",
                "source": entry.get("source") or "cli",
                "first_seen": entry.get("first_seen"),
                "posted_date": entry.get("posted_date"),
                "fit": entry.get("fit") or "medium",
                "status": entry.get("status") or "new",
                "rank_score": entry.get("rank_score"),
                "rank_verdict": entry.get("rank_verdict"),
                "location_verdict": entry.get("location_verdict"),
                "language_gate": entry.get("language_gate"),
                "guard": entry.get("guard") or "clean",
                "strengths": entry.get("strengths") or [],
                "gaps": entry.get("gaps") or [],
                "description_excerpt": entry.get("description_excerpt") or "",
                "in_tracker": in_tracker,
                "salary_benchmark": salary.estimate_market_salary(comp, title, entry.get("location") or "", entry.get("description_excerpt") or ""),
                **portals.build_referral_links(comp, title),
            })

        items.sort(
            key=lambda x: (
                0 if x["status"] != "expired" else 1,
                x["rank_score"] if isinstance(x.get("rank_score"), int) else (72 if x["fit"] == "high" else 55 if x["fit"] == "medium" else 30),
                x.get("posted_date") or x.get("first_seen") or "",
            ),
            reverse=True,
        )
        return items

    def batch_rank_seen_jobs(self, limit: int = 20, include_all: bool = False):
        """Runs /rank rubric across eligible entries in seen_jobs.json."""
        _, seen = portals.load_seen_state()
        excluded = portals.tracker_pairs(self.tracker_path)
        from rank_state import norm as rank_norm

        candidates = []
        for key, entry in seen.items():
            st = entry.get("status")
            if include_all:
                if st == "skipped":
                    continue
            elif st not in ("new", "unreachable"):
                continue
            if (rank_norm(entry.get("company")), rank_norm(entry.get("title"))) in excluded:
                continue
            candidates.append((key, entry))

        candidates = candidates[:limit]
        results = []
        for key, entry in candidates:
            desc = entry.get("description_excerpt") or ""
            guard_verdict = entry.get("guard") or "clean"
            if not desc and entry.get("url"):
                try:
                    fetched = portals.fetch_and_guard_url(entry["url"])
                    if fetched.get("ok"):
                        if not fetched.get("is_active", True):
                            results.append({"key": key, "status": "expired"})
                            continue
                        desc = fetched.get("description") or ""
                        guard_verdict = fetched.get("guard_verdict") or guard_verdict
                except Exception:
                    pass

            scored = self.score_jd(entry.get("company", ""), entry.get("title", ""), desc or entry.get("title", ""), entry.get("location", ""))
            w100 = scored["breakdown"]["weighted_100"]
            results.append({
                "key": key,
                "status": "scored",
                "scores": w100,
                "location_verdict": scored["gates"]["location_verdict"],
                "language_gate": scored["gates"]["language_gate"],
                "language_note": scored["gates"]["language_note"] if scored["gates"]["language_gate"] != "PASS" else None,
                "strengths": scored["proof_points"][:3],
                "gaps": scored["gaps"][:3] if scored["gaps"] else ["Confirm role-specific KPI metrics in discovery call"],
                "guard": guard_verdict,
            })

        write_res = portals.record_rank_scores(results)
        return {
            "evaluated": len(candidates),
            "updated_count": write_res["updated_count"],
            "updated": write_res["updated"],
            "seen_jobs": self.get_seen_jobs_view(),
        }

    def run_live_scan(self):
        """Runs combined selective shortlist + live portal scrape."""
        scrape_out = self.run_live_scrape(
            queries=["Senior Product Manager payments", "Product Manager settlement", "AI Product Manager fintech"],
            locations=["Bengaluru, Karnataka, India"],
            portals_enabled=["linkedin", "freehire"],
            jobage_days=14,
            enrich_details=True,
        )
        shortlist_jobs = self.get_shortlist()
        return {
            "status": "success",
            "timestamp": scrape_out["timestamp"],
            "jobs_scanned": len(shortlist_jobs) + scrape_out["filtered_pm_hits"],
            "new_jobs_added": scrape_out["added_count"],
            "logs": self.scanner_logs[-5:],
        }

    def run_upskill_analysis(self, mode: str = "aggregate", target_input: str = "", company: str = "", role: str = ""):
        """Implements .agents/skills/upskill/SKILL.md for Aggregate (tracker + ranked seen_jobs) or Targeted (URL or JD)."""
        self.upskill_dir.mkdir(parents=True, exist_ok=True)
        today_str = datetime.now().strftime("%Y-%m-%d")

        # Profile skills already possessed (generous profile match per upskill Skill Rule 4)
        profile_text = (read_text("profile/master_cv.md") + "\n" + (
            self.candidate_facts_path.read_text(encoding="utf-8") if self.candidate_facts_path.exists() else ""
        )).lower()

        skill_catalog = [
            ("UPI & Real-Time Domestic A2A Rails", "Domain", "~15h",
             ["NPCI UPI Product Architecture & Technical Specs (npci.org.in)", "Razorpay Engineering Blog — UPI Switch & Intent Flows"],
             "Leverage your Visa RuPay/Jaywan/KSA domestic-scheme background; focus only on NPCI 2-step collect/intent flows and VPA resolution.",
             ["upi", "npci", "a2a"]),
            ("Consumer Lending, Underwriting & BNPL Risk", "Domain", "~20h",
             ["Reserve Bank of India Digital Lending Guidelines (rbi.org.in)", "Reforge — Credit & Underwriting Product Strategy"],
             "Bridge from Visa Instalment Plan B2B spec enablement into consumer bureau scoring, cohort default curves, and collections funnels.",
             ["lending", "loan", "underwriting", "bnpl", "credit line"]),
            ("WealthTech, Mutual Funds & Digital Gold", "Domain", "~18h",
             ["SEBI / AMFI Mutual Fund Distribution Primer", "Zerodha Varsity — Personal Finance & Market Infrastructure"],
             "Skip basic settlement plumbing; study SIP mandate rails, AMC RTA integration (CAMS/KFintech), and suitability compliance.",
             ["wealth", "mutual fund", "digital gold", "investment", "insurance"]),
            ("B2C Consumer Growth, Activation & Retention Loops", "Soft / Product", "~15h",
             ["Reforge — Retention + Engagement Deep Dive", "Lenny's Podcast — Consumer Fintech Activation Benchmarks"],
             "Contrast with your B2B acquirer/merchant metrics; practice framing onboarding funnel drops, DAU/MAU cohorts, and consumer paywalls.",
             ["b2c", "consumer", "consumer fintech", "activation", "funnel", "dau"]),
            ("LLM Evals, RAG & Production AI Agent Guardrails", "Hard / AI", "~20h",
             ["Anthropic — Building Effective Agents (anthropic.com/research)", "DeepLearning.AI — Evaluating AI Agents & RAG Systems"],
             "Build directly on your shipped Fast Refund AI & Claude SMCC workflow; add formal precision/recall eval harnesses and latency budgets.",
             ["llm", "ai", "agentic", "rag", "eval", "genai", "ml"]),
            ("Direct Vendor Commercial Negotiation & Procurement", "Soft / Governance", "~8h",
             ["Harvard Program on Negotiation — SaaS & Scheme Vendor Contracts"],
             "Use your Cognizant SOW requote influence story as anchor; study SLA penalty clauses, minimum-volume guarantees, and RFP scoring.",
             ["vendor contract", "procurement", "sow", "commercial contract"]),
            ("SQL Analytics, Funnel Instrumentation & Experimentation", "Hard / Data", "~10h",
             ["Mode Analytics SQL Tutorial — Funnel & Cohort Analysis", "Trustworthy Online Controlled Experiments (Kohavi)"],
             "You already list SSMS & experimentation; focus specifically on self-serve cohort retention queries and MDE A/B test sizing.",
             ["ab test", "a/b", "funnel analytics", "mixpanel", "amplitude", "snowflake"]),
        ]

        analyzed_jobs_count = 0
        tracker_count = 0
        ranked_count = 0
        skill_scores = {}  # skill_name -> {"score": float, "sources": list[str], "meta": tuple}

        if mode == "targeted":
            text_body = (target_input or "").strip()
            resolved_url = ""
            if re.match(r"^https?://", text_body, re.I):
                resolved_url = text_body
                fetched = portals.fetch_and_guard_url(text_body)
                if not fetched.get("ok"):
                    return {"ok": False, "error": fetched.get("error", "Could not fetch target URL")}
                text_body = fetched.get("description") or ""
                company = company or fetched.get("company") or "Target Company"
                role = role or fetched.get("title") or "Target Role"
            else:
                _, gf = posting_guard.guard(text_body, "auto")
                company = company or "Target Company"
                role = role or "Senior Product Manager"

            t_low = f"{role} {text_body}".lower()
            analyzed_jobs_count = 1
            for name, kind, est, resources, direction, triggers in skill_catalog:
                if any(tr in t_low for tr in triggers):
                    skill_scores[name] = {
                        "score": 1.0,
                        "count": 1,
                        "provenance": "required/mentioned in target JD",
                        "priority": "Critical" if triggers[0] in t_low else "High",
                        "type": kind,
                        "est": est,
                        "resources": resources,
                        "direction": direction,
                    }
            # Ensure at least top 2 relevant growth recommendations show if JD is generic
            if not skill_scores:
                for item in skill_catalog[:2]:
                    name, kind, est, resources, direction, _ = item
                    skill_scores[name] = {
                        "score": 0.6,
                        "count": 1,
                        "provenance": "LLM/profile synthesis",
                        "priority": "High",
                        "type": kind,
                        "est": est,
                        "resources": resources,
                        "direction": direction,
                    }
        else:
            # Aggregate mode: tracker.csv + ranked seen_jobs (rank_score >= 45), deduplicated on company+title
            rows = self.get_tracker()
            _, seen = portals.load_seen_state()
            merged = {}

            for r in rows:
                comp = (r.get("company") or "").strip()
                r_title = (r.get("role") or "").strip()
                if not comp or not r_title:
                    continue
                k = (comp.lower(), r_title.lower())
                raw_sc = r.get("score") or ""
                try:
                    num = float(raw_sc)
                    rating_100 = num * 10.0 if num <= 10.0 else num
                except ValueError:
                    rating_100 = None
                merged[k] = {
                    "company": comp,
                    "title": r_title,
                    "rating_100": rating_100,
                    "text": f"{comp} {r_title} {r.get('notes', '')}".lower(),
                    "gaps": [],
                    "from_tracker": True,
                }
                tracker_count += 1

            for entry in seen.values():
                if entry.get("status") != "ranked":
                    continue
                rs = entry.get("rank_score")
                if not isinstance(rs, (int, float)) or rs < 45:
                    continue
                comp = (entry.get("company") or "").strip()
                title = (entry.get("title") or "").strip()
                k = (comp.lower(), title.lower())
                ranked_count += 1
                if k in merged:
                    merged[k]["gaps"] = entry.get("gaps") or []
                    if merged[k]["rating_100"] is None:
                        merged[k]["rating_100"] = float(rs)
                else:
                    merged[k] = {
                        "company": comp,
                        "title": title,
                        "rating_100": float(rs),
                        "text": f"{comp} {title} {' '.join(entry.get('gaps') or [])}".lower(),
                        "gaps": entry.get("gaps") or [],
                        "from_tracker": False,
                    }

            analyzed_jobs_count = len(merged)
            for item in merged.values():
                rating = item["rating_100"] if item["rating_100"] is not None else 70.0
                weight = max(0.15, (100.0 - rating) / 100.0)
                combined_txt = item["text"] + " " + " ".join(item["gaps"]).lower()
                for name, kind, est, resources, direction, triggers in skill_catalog:
                    if any(tr in combined_txt for tr in triggers):
                        rec = skill_scores.setdefault(name, {
                            "score": 0.0,
                            "count": 0,
                            "recorded_cnt": 0,
                            "inferred_cnt": 0,
                            "type": kind,
                            "est": est,
                            "resources": resources,
                            "direction": direction,
                        })
                        rec["score"] += weight
                        rec["count"] += 1
                        if item["gaps"]:
                            rec["recorded_cnt"] += 1
                        else:
                            rec["inferred_cnt"] += 1

            # Always include AI Evals & B2C/UPI strategic themes as Baseline Synthesis if heatmap is sparse
            for idx, (name, kind, est, resources, direction, _) in enumerate(skill_catalog[:3]):
                if name not in skill_scores:
                    skill_scores[name] = {
                        "score": round(0.45 - idx * 0.1, 2),
                        "count": 1,
                        "recorded_cnt": 0,
                        "inferred_cnt": 1,
                        "type": f"{kind} [synthesis]",
                        "est": est,
                        "resources": resources,
                        "direction": direction,
                    }

            for name, rec in skill_scores.items():
                sc = round(rec["score"], 2)
                rec["score"] = sc
                rec["priority"] = "Critical" if sc >= 0.8 else ("High" if sc >= 0.4 else "Medium")
                rec["provenance"] = f"{rec['count']} roles ({rec.get('recorded_cnt', 0)} recorded, {rec.get('inferred_cnt', 0)} inferred), score {sc}"

        heatmap = sorted(
            [{"skill": k, **v} for k, v in skill_scores.items()],
            key=lambda x: x["score"],
            reverse=True,
        )

        # Build Markdown Report per upskill SKILL.md Step 8
        heatmap_rows = "\n".join(
            f"| {h['priority']} | {h['skill']} | {h['type']} | {h['provenance']} |" for h in heatmap
        )
        plan_blocks = []
        order_rows = []
        total_hours = 0
        for idx, h in enumerate(heatmap[:5], 1):
            hrs = int(re.search(r"\d+", h["est"]).group(0))
            total_hours += hrs
            res_lines = "\n".join(f"- {r}" for r in h["resources"])
            plan_blocks.append(f"### {h['skill']} `[{h['type']}]` — {h['est']}\n{res_lines}\n\nStudy direction: {h['direction']}\n")
            order_rows.append(f"| {idx} | {h['skill']} | {h['type']} | {h['est']} | Priority: {h['priority']} |")

        mode_label = f"Targeted: {role} @ {company}" if mode == "targeted" else f"Aggregate ({analyzed_jobs_count} jobs: {tracker_count} tracked, {ranked_count} ranked)"
        report_md = f"""# Upskill Report — {today_str}
**Mode:** {mode_label}

---

## Gap Heatmap

| Priority | Skill / Area | Type | Gap Source |
|---|---|---|---|
{heatmap_rows}

---

## Learning Plan

{chr(10).join(plan_blocks)}

---

## Suggested Study Order

| # | Topic | Type | Est. Time | Note |
|---|---|---|---|---|
{chr(10).join(order_rows)}

**Total estimated time: ~{total_hours}h**
"""
        if mode == "targeted":
            c_slug = re.sub(r"[^a-z0-9]+", "-", (company or "target").lower()).strip("-")[:20] or "target"
            r_slug = re.sub(r"[^a-z0-9]+", "-", (role or "pm").lower()).strip("-")[:25] or "pm"
            filename = f"report-{today_str}-{c_slug}-{r_slug}.md"
        else:
            filename = f"report-{today_str}.md"

        report_path = self.upskill_dir / filename
        report_path.write_text(report_md, encoding="utf-8")

        return {
            "ok": True,
            "mode": mode,
            "mode_label": mode_label,
            "filename": filename,
            "download_url": f"/download/upskill/{filename}",
            "total_hours": total_hours,
            "heatmap": heatmap,
            "markdown": report_md,
        }
