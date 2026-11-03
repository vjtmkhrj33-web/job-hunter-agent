---
name: fit-scorer
description: Scores a job description against the master CV on a 1-10 rubric and lists strengths, gaps and a go/no-go. Use when given a job URL or pasted JD, or when triaging jobs/shortlist.md.
---
# Fit scoring

Input: a job link or pasted JD. Save the JD to applications/<company>-<role>/jd.md.
Score 1-10 using these weights:
- Domain match, payments / acquiring / settlement / platform (30%)
- Seniority and scope match (20%)
- Technical depth match, APIs, event-driven, specs (15%)
- Evidence of measurable impact relevant to the role (15%)
- Location / sponsorship / compensation fit per targets.md (20%)

Output (applications/<company>-<role>/fit.md):
- Score and one-line verdict (APPLY / APPLY WITH REFERRAL ONLY / SKIP)
- Top 5 matching proof points from the master CV (quote the CV wording)
- Gaps, stated honestly, with how to address them in the interview (never hide them in the resume)
- Keywords from the JD that truthfully appear in my profile
Only recommend APPLY if score is 7 or higher.
