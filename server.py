import os
import sys
import json
import threading
import time
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, unquote

sys.path.insert(0, str(Path(__file__).resolve().parent))
from core.engine import JobAgentEngine
from core.integrations import IntegrationsManager
from core import portals

BASE_DIR = Path(__file__).resolve().parent
engine = JobAgentEngine()
integrations_mgr = IntegrationsManager()


def background_scanner_worker():
    while True:
        if engine.scanner_active:
            try:
                engine.run_live_scan()
            except Exception as e:
                engine.scanner_logs.append(f"[Error] Scanner error: {str(e)}")
        time.sleep(10800)


scan_thread = threading.Thread(target=background_scanner_worker, daemon=True)
scan_thread.start()


class JobAgentHTTPHandler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/" or path == "/index.html":
            ui_path = BASE_DIR / "ui" / "index.html"
            if ui_path.exists():
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                with open(ui_path, "rb") as f:
                    self.wfile.write(f.read())
                return
            self.send_error(404, "UI not found")
            return

        if path.startswith("/download/"):
            rel_path = unquote(path[len("/download/"):]).lstrip("/\\")
            target_file = (BASE_DIR / rel_path).resolve()
            try:
                target_file.relative_to(BASE_DIR.resolve())
            except ValueError:
                self.send_error(403, "Invalid file path")
                return
            if target_file.exists() and target_file.is_file():
                self.send_response(200)
                content_type = (
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    if target_file.suffix == ".docx"
                    else "text/plain; charset=utf-8"
                )
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Disposition", f'attachment; filename="{target_file.name}"')
                self.send_header("Content-Length", str(target_file.stat().st_size))
                self.end_headers()
                with open(target_file, "rb") as f:
                    self.wfile.write(f.read())
                return
            self.send_error(404, f"File not found: {rel_path}")
            return

        # API Endpoints
        if path == "/api/stats":
            self.send_json(engine.get_stats())
        elif path == "/api/tracker":
            self.send_json(engine.get_tracker())
        elif path == "/api/shortlist":
            self.send_json(engine.get_shortlist())
        elif path == "/api/scrape/seen":
            self.send_json({"jobs": engine.get_seen_jobs_view()})
        elif path == "/api/skills/status":
            self.send_json(engine.get_skills_overview())
        elif path == "/api/references":
            self.send_json(engine.get_references_data())
        elif path == "/api/integrations":
            self.send_json(self.get_integrations_status())
        elif path == "/api/integrations/config":
            self.send_json(integrations_mgr.config)
        elif path == "/api/variants":
            self.send_json(self.get_cv_variants())
        elif path == "/api/prep/brief":
            query = parse_qs(parsed.query)
            company = query.get("company", [""])[0]
            role = query.get("role", [""])[0]
            self.send_json(self.get_prep_brief(company, role))
        elif path == "/api/prep/deepdive":
            query = parse_qs(parsed.query)
            company = query.get("company", [""])[0]
            role = query.get("role", [""])[0]
            self.send_json(engine.get_interview_prep_deep_dive(company, role))
        elif path == "/api/applications/file":
            query = parse_qs(parsed.query)
            slug = query.get("slug", [""])[0].strip("/")
            name = query.get("name", [""])[0].strip("/")
            target = (engine.applications_dir / slug / name).resolve()
            try:
                target.relative_to(engine.applications_dir.resolve())
                content = target.read_text(encoding="utf-8") if target.is_file() else ""
                self.send_json({"ok": target.is_file(), "filename": name, "content": content})
            except Exception:
                self.send_json({"ok": False, "error": "Invalid file path"})
        elif path.startswith("/api/applications"):
            self.send_json(self.get_applications_list())
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8") if length > 0 else "{}"
        try:
            data = json.loads(body)
        except Exception:
            data = {}

        if path in ("/api/tracker/update", "/api/tracker/add"):
            company = data.get("company", "")
            role = data.get("role", "")
            status = data.get("status", "shortlisted")
            notes = data.get("notes", "")
            follow_up = data.get("follow_up_date", "")
            link = data.get("link", "")
            location = data.get("location", "")
            sponsorship = data.get("sponsorship", "")
            score = data.get("score", "")
            engine.update_tracker_status(
                company, role, status, notes, follow_up,
                link=link, location=location, sponsorship=sponsorship, score=score,
            )
            self.send_json({"status": "success", "tracker": engine.get_tracker()})

        elif path == "/api/jobs/fetch_url":
            url = data.get("url", "")
            res = portals.fetch_and_guard_url(url)
            self.send_json(res)

        elif path == "/api/jobs/score":
            company = data.get("company", "Target Company")
            role = data.get("role", "Senior Product Manager")
            jd_text = data.get("jd_text", "")
            location = data.get("location", "Bengaluru")
            result = engine.score_jd(company, role, jd_text, location)
            self.send_json(result)

        elif path == "/api/jobs/tailor":
            company = data.get("company", "Target Company")
            role = data.get("role", "Senior Product Manager")
            jd_text = data.get("jd_text", "")
            location = data.get("location", "Bengaluru")
            job_url = data.get("job_url", "")
            result = engine.tailor_application(company, role, jd_text, location, job_url=job_url)
            self.send_json(result)

        elif path == "/api/scrape/run":
            queries = data.get("queries")
            locations = data.get("locations")
            portals_enabled = data.get("portals")
            jobage_days = int(data.get("jobage_days") or 14)
            remote = data.get("remote") or None
            enrich_details = bool(data.get("enrich_details", True))
            res = engine.run_live_scrape(
                queries=queries,
                locations=locations,
                portals_enabled=portals_enabled,
                jobage_days=jobage_days,
                remote=remote,
                enrich_details=enrich_details,
                integrations_mgr=integrations_mgr,
            )
            self.send_json(res)

        elif path == "/api/scrape/rank":
            limit = int(data.get("limit") or 20)
            include_all = bool(data.get("include_all", False))
            res = engine.batch_rank_seen_jobs(limit=limit, include_all=include_all)
            self.send_json(res)

        elif path == "/api/upskill/run":
            mode = data.get("mode", "aggregate")
            target_input = data.get("target_input", "")
            company = data.get("company", "")
            role = data.get("role", "")
            res = engine.run_upskill_analysis(mode=mode, target_input=target_input, company=company, role=role)
            self.send_json(res)

        elif path == "/api/references/fact":
            category = data.get("category", "Updates")
            fact = data.get("fact", "")
            res = engine.append_candidate_fact(category, fact)
            self.send_json(res)

        elif path == "/api/scanner/run":
            res = engine.run_live_scan()
            self.send_json(res)

        elif path == "/api/scanner/toggle":
            engine.scanner_active = not engine.scanner_active
            self.send_json({"scanner_active": engine.scanner_active})

        elif path == "/api/prep/mock":
            question = data.get("question", "")
            answer = data.get("answer", "")
            evaluation = engine.evaluate_mock_answer(question, answer)
            self.send_json(evaluation)

        elif path == "/api/followup":
            company = data.get("company", "")
            role = data.get("role", "")
            applied_date = data.get("applied_date", "recent")
            days_elapsed = data.get("days_elapsed", "7")
            contact_name = data.get("contact_name", "Hiring Team")
            res = engine.generate_follow_up_pack(company, role, applied_date, days_elapsed, contact_name)
            self.send_json(res)

        elif path == "/api/integrations/config":
            res = integrations_mgr.save_config(data)
            self.send_json(res)

        elif path == "/api/integrations/test_gmail":
            email_val = data.get("email", "")
            pwd_val = data.get("app_password", "")
            res = integrations_mgr.test_gmail_connection(email_val, pwd_val)
            self.send_json(res)

        elif path == "/api/integrations/scan_gmail":
            res = integrations_mgr.scan_gmail_for_interviews()
            self.send_json(res)

        elif path == "/api/integrations/check_calendar":
            res = integrations_mgr.check_calendar_feed()
            self.send_json(res)

        elif path == "/api/crawler/live_crawl":
            jobs = integrations_mgr.crawl_public_job_boards()
            scored = []
            for j in jobs:
                fit = engine.score_jd(j["company"], j["title"], j["title"] + " " + j["location"], j["location"])
                if fit["score"] >= 7.0:
                    scored.append({**j, "fit": fit})
            self.send_json({"total_found": len(jobs), "high_fit_jobs": scored})

        else:
            self.send_error(404, "Endpoint not found")

    def send_json(self, data):
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data, indent=2).encode("utf-8"))

    def get_applications_list(self):
        apps = []
        if engine.applications_dir.exists():
            for d in sorted(engine.applications_dir.iterdir(), key=lambda p: p.stat().st_mtime, reverse=True):
                if d.is_dir():
                    files = [f.name for f in d.iterdir()]
                    docx_files = sorted([f for f in files if f.endswith(".docx")], reverse=True)
                    apps.append({
                        "slug": d.name,
                        "files": files,
                        "docx": docx_files[0] if docx_files else None,
                        "has_resume": "resume.md" in files,
                        "has_cover_note": "cover_note.md" in files,
                        "has_fit": "fit.md" in files,
                        "has_outreach": "outreach.md" in files,
                        "has_answers": "application_answers.md" in files,
                        "has_quality_check": "pre_submit_check.md" in files,
                    })
        return apps

    def get_cv_variants(self):
        return [
            {
                "id": "payments_ai",
                "title": "Payments, Acquiring & Core Rails Specialist",
                "target_roles": "Senior Payments PM (I6), Acquiring Platform PM, Core Settlement Lead",
                "docx": "/download/profile/master_cv_payments_ai.docx",
                "md": "/download/profile/master_cv_payments_ai.md",
                "description": "Emphasizes merchant acquiring, ISO 8583, BASE I/II, domestic schemes (UAE Jaywan, RuPay, KSA), Fast Refund AI, and Kafka settlement.",
            },
            {
                "id": "leadership",
                "title": "Product Leadership & Stretch PM (Group / Director)",
                "target_roles": "Group Product Manager, Director of PM (Stretch Application)",
                "docx": "/download/profile/master_cv_leadership.docx",
                "md": "/download/profile/master_cv_leadership.md",
                "description": "Leads with Cognizant 12-person management, 100+ interviews, RICE VP roadmap sign-off, and $30M net revenue impact.",
            },
            {
                "id": "generalist",
                "title": "Generalist Senior Technical Product Manager",
                "target_roles": "Senior Product Manager, Technical PM, Platform PM",
                "docx": "/download/profile/master_cv_generalist.docx",
                "md": "/download/profile/master_cv_generalist.md",
                "description": "Balanced view of 0->1 discovery, backlog prioritization, CSPO delivery across 4 scrum teams, and microservices modernization.",
            },
        ]

    def get_prep_brief(self, company, role):
        company_clean = company or "Target Company"
        questions = [
            {
                "id": "q1",
                "category": "Product Sense & Architecture",
                "question": f"How would you approach modernizing a legacy payment processing platform to support event-driven settlement at {company_clean}?",
                "star_hook": "Payments 2.0 domestic scheme architecture ($30M net revenue) & Kafka settlement events.",
            },
            {
                "id": "q2",
                "category": "Operational Efficiency & Automation",
                "question": "Tell me about a time you identified a massive recurring operational bottleneck and automated it.",
                "star_hook": "Visa CAS automation freed 10 FTEs (<24h turnaround); Cognizant Transaction Validation System freed ~30 analysts into Project Lead roles.",
            },
            {
                "id": "q3",
                "category": "Crisis / Escalation Under Pressure",
                "question": "How do you handle severe P1 production failures or SWAT escalations that threaten merchant revenue?",
                "star_hook": "$9.2M Visa protection (attributed to Customer Support post-release), KFC $1.5M auth-to-settlement recapture, SME SWAT & P1 call lead.",
            },
            {
                "id": "q4",
                "category": "AI in Regulated Payments",
                "question": "How have you shipped practical AI improvements in regulated payments without introducing risk?",
                "star_hook": "Fast Refund AI removed unnecessary spike analysis/regression (24h refunds, 97% success); SMCC Japan Claude specs cut 6 wks to 2 wks.",
            },
            {
                "id": "q5",
                "category": "Regulatory Trade-Off & Influence Without Authority",
                "question": "Tell me about a difficult trade-off or a time you drove change outside your direct authority.",
                "star_hook": "Apple Pay India data center move (RBI compliance over negligible cost); Acquirer Onboarding UI change with architect approval; Cognizant SOW requote.",
            },
        ]
        return {
            "company": company_clean,
            "role": role or "Senior Product Manager",
            "questions": questions,
        }

    def get_integrations_status(self):
        return {
            "linkedin": {
                "name": "LinkedIn Public Guest Search",
                "status": "Active (Safe Unauthenticated Mode)",
                "description": "Queries public LinkedIn jobs-guest endpoints & runs posting_guard sanitization.",
            },
            "freehire": {
                "name": "Freehire ATS Aggregator API",
                "status": "Active (~50 ATS Platforms)",
                "description": "Searches live Greenhouse, Lever, Ashby, and employer ATS postings via freehire.me.",
            },
            "gmail": {
                "name": "Gmail Inbox & Thread Tracker",
                "status": "Local IMAP Ready",
                "description": "Scans for interview invites and recruiter responses.",
            },
            "calendar": {
                "name": "Google Calendar Sync",
                "status": "iCal Ready",
                "description": "Monitors interview calendar events.",
            },
        }


def run_server(port=8080):
    server_address = ("", port)
    httpd = HTTPServer(server_address, JobAgentHTTPHandler)
    print("======================================================")
    print("  Job Hunter & Application Tracker Dashboard running!")
    print(f"  Access UI at: http://localhost:{port}")
    print("======================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down server...")
        httpd.server_close()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    run_server(port)
