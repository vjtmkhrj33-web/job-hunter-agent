---
name: resume-tailor
description: Produces a one-page, ATS-safe tailored resume and short cover note for a specific job from the master CV, without inventing anything. Use after a job is marked APPLY.
---
# Resume tailoring

Input: applications/<company>-<role>/jd.md and fit.md.
1. Reorder and trim master CV content to the 6-8 bullets most relevant to this JD. Mirror the JD's language only where it truthfully describes my work.
2. Rewrite the headline and summary (3 lines max) for this role.
3. Keep metrics exactly as in the master CV, with the same claim strength. If two metrics conflict, use the more conservative one and flag it to me.
4. Output resume.md (plain text, ATS-safe, one page) and cover_note.md (under 150 words: why this company, 2 proof points, one clear ask).
5. End with a "Verify before sending" list of any line you were unsure about.
Never add skills, tools or domains not in the master CV.
