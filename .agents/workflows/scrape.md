Run the `job-scraper` skill (`.agents/skills/job-scraper/SKILL.md`) using `.agents/skills/job-scraper/search-queries.md` and `profile/targets.md`:
1. Check `bun --version` and run enabled portal CLIs (`linkedin-search`, `freehire-search`) in `.agents/skills/` for roles posted in the last 14 days across Bengaluru and sponsorship-verified target markets.
2. Deduplicate candidates with `python tools/scrape_state.py candidates --results <json>` against `job_scraper/seen_jobs.json` and `tracker.csv`.
3. Fetch details for new matches, pass fetched posting text through `python tools/posting_guard.py` before reading, store results via `python tools/scrape_state.py store --results <json>`, and present a prioritized shortlist with LinkedIn referral/peer search links. Do not apply or send messages.

