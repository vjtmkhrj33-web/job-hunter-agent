Batch-score newly scraped jobs in `job_scraper/seen_jobs.json`:
1. Select eligible candidates with `python tools/rank_state.py candidates --limit 10` (reads `tracker.csv` automatically to exclude already-tracked roles).
2. Fetch any missing posting bodies using the escalation order in `.agents/skills/job-application-assistant/09-web-research.md` (`linkedin-search detail` / `freehire-search detail`, `WebFetch`, `python tools/fetch_page.py`) and gate every posting with `python tools/posting_guard.py`.
3. Score each role against `profile/master_cv.md`, `.agents/skills/senior-pm-job-search/references/candidate-facts.md`, `.agents/skills/senior-pm-job-search/references/search-profile.md`, and `.agents/skills/job-application-assistant/04-job-evaluation.md` (including Eligibility & Language gates and visa sponsorship rules).
4. Persist scores with `python tools/rank_state.py apply --results <json>` and present the ranked shortlist with red flags and closing-soon deadlines.

