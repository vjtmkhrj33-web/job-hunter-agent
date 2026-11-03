import os
import sys
import json
import threading
import time
from pathlib import Path
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

sys.path.insert(0, str(Path(__file__).resolve().parent))
from core.engine import JobAgentEngine
from core.integrations import IntegrationsManager

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
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
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
            else:
                self.send_error(404, "UI not found")
                return

        if path.startswith("/download/"):
            rel_path = path.replace("/download/", "")
            target_file = BASE_DIR / rel_path
            if target_file.exists() and target_file.is_file():
                self.send_response(200)
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if target_file.suffix == ".docx" else "text/plain; charset=utf-8"
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Disposition", f'attachment; filename="{target_file.name}"')
                self.send_header("Content-Length", str(target_file.stat().st_size))
                self.end_headers()
                with open(target_file, "rb") as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_error(404, f"File not found: {rel_path}")
                return

        # API Endpoints
        if path == "/api/stats":
            self.send_json(engine.get_stats())
        elif path == "/api/tracker":
            self.send_json(engine.get_tracker())
        elif path == "/api/shortlist":
            self.send_json(engine.get_shortlist())
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
        elif path.startswith("/api/applications"):
            self.send_json(self.get_applications_list())
        else:
            super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length).decode('utf-8') if length > 0 else "{}"
        try:
            data = json.loads(body)
        except Exception:
            data = {}

        if path == "/api/tracker/update":
            company = data.get("company", "")
            role = data.get("role", "")
            status = data.get("status", "")
            notes = data.get("notes", "")
            follow_up = data.get("follow_up_date", "")
            engine.update_tracker_status(company, role, status, notes, follow_up)
            self.send_json({"status": "success", "tracker": engine.get_tracker()})

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
            result = engine.tailor_application(company, role, jd_text, location)
            self.send_json(result)

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
            # Score each job
            scored = []
            for j in jobs:
                fit = engine.score_jd(j['company'], j['title'], j['title'] + " " + j['location'], j['location'])
                if fit['score'] >= 7.0:
                    scored.append({**j, "fit": fit})
            self.send_json({"total_found": len(jobs), "high_fit_jobs": scored})

        else:
            self.send_error(404, "Endpoint not found")

    def send_json(self, data):
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data, indent=2).encode('utf-8'))

    def get_applications_list(self):
        apps = []
        if engine.applications_dir.exists():
            for d in engine.applications_dir.iterdir():
                if d.is_dir():
                    files = [f.name for f in d.iterdir()]
                    docx_files = [f for f in files if f.endswith(".docx")]
                    apps.append({
                        "slug": d.name,
                        "files": files,
                        "docx": docx_files[0] if docx_files else None,
                        "has_resume": "resume.md" in files,
                        "has_cover_note": "cover_note.md" in files,
                        "has_fit": "fit.md" in files,
                        "has_outreach": "outreach.md" in files,
                        "has_answers": "application_answers.md" in files,
                        "has_quality_check": "pre_submit_check.md" in files
                    })
        return apps

    def get_cv_variants(self):
        return [
            {
                "id": "payments_ai",
                "title": "Payments, Acquiring & Core Rails Specialist",
                "target_roles": "Senior Payments PM, Acquiring Platform PM, Core Settlement Lead",
                "docx": "/download/profile/master_cv_payments_ai.docx",
                "md": "/download/profile/master_cv_payments_ai.md",
                "description": "Emphasizes merchant acquiring, ISO 8583, BASE I/II, domestic schemes (UAE Jaywan, RuPay, KSA), and Kafka event-driven settlement."
            },
            {
                "id": "leadership",
                "title": "Product Leadership & Group Product Manager (GPM)",
                "target_roles": "Group Product Manager, Head of Platform Delivery, Director of PM",
                "docx": "/download/profile/master_cv_leadership.docx",
                "md": "/download/profile/master_cv_leadership.md",
                "description": "Leads with people management (12 PMs), P&L revenue impact ($30M generated, $22M+ protected), and enterprise cross-functional governance."
            },
            {
                "id": "generalist",
                "title": "Generalist Senior Technical Product Manager",
                "target_roles": "Senior Product Manager, Technical PM, Platform PM",
                "docx": "/download/profile/master_cv_generalist.docx",
                "md": "/download/profile/master_cv_generalist.md",
                "description": "Balanced view of 0→1 discovery, backlog prioritization, CSPO delivery, microservices architecture, and AI-accelerated workflows."
            }
        ]

    def get_prep_brief(self, company, role):
        company_clean = company or "Target Company"
        questions = [
            {
                "id": "q1",
                "category": "Product Sense & Architecture",
                "question": f"How would you approach modernizing a legacy payment processing platform to support event-driven settlement at {company_clean}?",
                "star_hook": "Payments 2.0 domestic scheme architecture ($30M net revenue) & Kafka settlement events."
            },
            {
                "id": "q2",
                "category": "Operational Efficiency & Automation",
                "question": "Tell me about a time you identified a massive recurring operational bottleneck in payment delivery and automated it.",
                "star_hook": "CAS self-service testing & settlement file retrieval (saved 8 FTEs/sprint, cut turnaround to <24 hrs)."
            },
            {
                "id": "q3",
                "category": "Crisis / Incident Management",
                "question": "How do you handle severe P1 production failures that directly threaten merchant revenue and settlement?",
                "star_hook": "$9.2M protected at Visa via automated regeneration & $11M protected at Fiserv via upstream validation."
            },
            {
                "id": "q4",
                "category": "AI in Product Management",
                "question": "CBD / modern fintechs are pushing for AI industrialization. How have you used AI agents in actual production product cycles?",
                "star_hook": "SMCC Japan integration delivered in 2 weeks instead of 6 using Claude agentic toolchain; cut PRD cycles by 50%."
            },
            {
                "id": "q5",
                "category": "Merchant & Acquirer Scale",
                "question": "How do you optimize merchant card conversion and tokenization across international markets?",
                "star_hook": "+25% token volume lift in Europe (COF/MIT/CIT) and +18% Visa Instalment Plan volume in CEMEA."
            }
        ]
        return {
            "company": company_clean,
            "role": role or "Senior Product Manager",
            "questions": questions
        }

    def get_integrations_status(self):
        return {
            "gmail": {
                "name": "Gmail Inbox & Thread Tracker",
                "status": "Ready for OAuth Setup",
                "description": "Automatically scans for interview invites, recruiter responses, and application receipts to update status in tracker.csv.",
                "permissions_required": ["gmail.readonly"],
                "local_port": 8080
            },
            "calendar": {
                "name": "Google Calendar Sync",
                "status": "Ready for OAuth Setup",
                "description": "Syncs scheduled interview rounds directly to your calendar with pre-interview STAR prep prompts 1 hour prior.",
                "permissions_required": ["calendar.events"],
                "local_port": 8080
            },
            "linkedin": {
                "name": "LinkedIn Public Scanner & Outreach",
                "status": "Connected (Safe Public Mode)",
                "description": "Discovers public job postings without risking account flags or scraping suspensions.",
                "permissions_required": ["None (Safe public query mode)"],
                "local_port": 8080
            },
            "drive": {
                "name": "Google Drive / Local Storage Sync",
                "status": "Active (Local Storage)",
                "description": "Local offline storage active in workspace profile and applications directory. Optional Google Drive sync.",
                "permissions_required": ["drive.file (Optional)"],
                "local_port": 8080
            }
        }

def run_server(port=8080):
    server_address = ('', port)
    httpd = HTTPServer(server_address, JobAgentHTTPHandler)
    print(f"======================================================")
    print(f"  Job Hunter & Application Tracker Dashboard running!")
    print(f"  Access local UI at: http://localhost:{port}")
    print(f"======================================================")
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
