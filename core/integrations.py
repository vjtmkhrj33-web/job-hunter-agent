import os
import re
import csv
import json
import imaplib
import email
from email.header import decode_header
import urllib.request
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path
import shutil

BASE_DIR = Path(__file__).resolve().parent.parent

class IntegrationsManager:
    def __init__(self):
        self.config_file = BASE_DIR / "config" / "integrations.json"
        self.config = self.load_config()

    def load_config(self):
        if self.config_file.exists():
            try:
                with open(self.config_file, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def save_config(self, new_config):
        self.config = new_config
        self.config_file.parent.mkdir(parents=True, exist_ok=True)
        with open(self.config_file, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)
        return {"status": "saved"}

    def test_gmail_connection(self, email_user, app_password):
        if not email_user or not app_password:
            return {
                "success": False,
                "message": "Email and 16-character App Password are required."
            }
        try:
            clean_pwd = app_password.replace(" ", "").strip()
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(email_user.strip(), clean_pwd)
            mail.select("inbox")
            status, messages = mail.search(None, 'ALL')
            mail.logout()
            return {
                "success": True,
                "message": f"Successfully connected to Gmail! Found inbox access for {email_user}."
            }
        except imaplib.IMAP4.error as e:
            return {
                "success": False,
                "message": f"Gmail Authentication failed: {str(e)}. Ensure you generate a 16-character Google App Password (not your normal password) at myaccount.google.com/apppasswords."
            }
        except Exception as e:
            return {
                "success": False,
                "message": f"Connection error: {str(e)}"
            }

    def scan_gmail_for_interviews(self):
        gmail_cfg = self.config.get("gmail", {})
        if not gmail_cfg.get("enabled"):
            return {"status": "disabled", "interviews_found": []}

        email_user = gmail_cfg.get("email", "")
        app_password = gmail_cfg.get("app_password", "").replace(" ", "").strip()

        if not email_user or not app_password:
            return {"status": "error", "message": "Gmail credentials not configured."}

        found_events = []
        try:
            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(email_user, app_password)
            mail.select("inbox")

            # Search recent emails (last 14 days)
            since_date = (datetime.now() - timedelta(days=14)).strftime("%d-%b-%Y")
            keywords = gmail_cfg.get("keywords", ["interview", "invitation", "recruiter", "next round"])

            for kw in keywords:
                status, data = mail.search(None, f'(SINCE "{since_date}" SUBJECT "{kw}")')
                if status == 'OK' and data[0]:
                    mail_ids = data[0].split()
                    for m_id in mail_ids[-5:]: # Check last 5 matching emails
                        status, msg_data = mail.fetch(m_id, '(RFC822)')
                        if status == 'OK':
                            raw_email = msg_data[0][1]
                            msg = email.message_from_bytes(raw_email)
                            
                            # Decode subject
                            subject, encoding = decode_header(msg["Subject"])[0]
                            if isinstance(subject, bytes):
                                subject = subject.decode(encoding if encoding else "utf-8", errors="ignore")
                            
                            from_sender = msg.get("From", "")
                            date_str = msg.get("Date", "")
                            
                            found_events.append({
                                "subject": subject,
                                "sender": from_sender,
                                "date": date_str,
                                "keyword": kw
                            })

            mail.logout()
            return {"status": "success", "interviews_found": found_events}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def check_calendar_feed(self):
        cal_cfg = self.config.get("calendar", {})
        if not cal_cfg.get("enabled") or not cal_cfg.get("ical_feed_url"):
            return {"status": "disabled", "upcoming_events": []}

        feed_url = cal_cfg.get("ical_feed_url", "").strip()
        try:
            req = urllib.request.Request(feed_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=10) as resp:
                content = resp.read().decode("utf-8", errors="ignore")

            # Simple iCal parser for VEVENT
            events = []
            vevents = re.findall(r'BEGIN:VEVENT(.*?)END:VEVENT', content, re.DOTALL)
            for v in vevents:
                summary_match = re.search(r'SUMMARY:(.*?)\r?\n', v)
                dtstart_match = re.search(r'DTSTART.*?:(.*?)\r?\n', v)
                summary = summary_match.group(1).strip() if summary_match else "Interview Event"
                dtstart = dtstart_match.group(1).strip() if dtstart_match else ""

                if any(k in summary.lower() for k in ["interview", "screening", "round", "avijit", "product"]):
                    events.append({
                        "summary": summary,
                        "time": dtstart
                    })

            return {"status": "success", "upcoming_events": events}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def mirror_to_drive(self, source_dir):
        drive_cfg = self.config.get("drive_storage", {})
        mirror_folder = drive_cfg.get("google_drive_mirror_folder", "").strip()

        if not mirror_folder or not os.path.exists(mirror_folder):
            return {"status": "no_mirror_path"}

        try:
            target = Path(mirror_folder) / Path(source_dir).name
            target.mkdir(parents=True, exist_ok=True)
            for file in Path(source_dir).glob("*"):
                if file.is_file():
                    shutil.copy2(file, target / file.name)
            return {"status": "success", "mirrored_to": str(target)}
        except Exception as e:
            return {"status": "error", "message": str(e)}

    def crawl_public_job_boards(self, target_companies=None):
        """Actively queries public Greenhouse & Lever boards for live Senior PM openings"""
        if not target_companies:
            target_companies = self.config.get("linkedin_and_web_crawler", {}).get("target_companies", [
                "Razorpay", "Adyen", "Stripe", "Airwallex", "Worldpay", "Checkout.com"
            ])

        found_jobs = []
        company_board_slugs = {
            "razorpay": ("greenhouse", "razorpaysoftwareprivatelimited"),
            "stripe": ("greenhouse", "stripe"),
            "airwallex": ("lever", "airwallex"),
            "checkout": ("lever", "checkout"),
            "adyen": ("custom", "https://careers.adyen.com")
        }

        for comp_name in target_companies:
            key = comp_name.lower().split()[0]
            if key in company_board_slugs:
                b_type, slug = company_board_slugs[key]
                if b_type == "greenhouse":
                    try:
                        url = f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs"
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=8) as resp:
                            data = json.loads(resp.read().decode("utf-8"))
                            jobs = data.get("jobs", [])
                            for j in jobs:
                                title = j.get("title", "")
                                loc = j.get("location", {}).get("name", "Bengaluru / Global")
                                if any(t.lower() in title.lower() for t in ["product manager", "product lead"]):
                                    found_jobs.append({
                                        "company": comp_name,
                                        "title": title,
                                        "location": loc,
                                        "link": j.get("absolute_url", ""),
                                        "source": "Greenhouse Live API",
                                        "posted_date": datetime.now().strftime("%Y-%m-%d")
                                    })
                    except Exception:
                        pass
                elif b_type == "lever":
                    try:
                        url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=8) as resp:
                            data = json.loads(resp.read().decode("utf-8"))
                            for j in data:
                                title = j.get("text", "")
                                loc = j.get("categories", {}).get("location", "APAC / Global")
                                if any(t.lower() in title.lower() for t in ["product manager", "product lead"]):
                                    found_jobs.append({
                                        "company": comp_name,
                                        "title": title,
                                        "location": loc,
                                        "link": j.get("hostedUrl", ""),
                                        "source": "Lever Live API",
                                        "posted_date": datetime.now().strftime("%Y-%m-%d")
                                    })
                    except Exception:
                        pass

        return found_jobs

