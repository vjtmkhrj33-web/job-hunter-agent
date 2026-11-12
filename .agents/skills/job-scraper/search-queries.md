# Search Queries for Job Scraper

Configured from `profile/targets.md` and `.agents/skills/senior-pm-job-search/references/search-profile.md`.

## Installed portal CLIs (primary for `/scrape`)

`/scrape` discovers every portal skill under `.agents/skills/*/SKILL.md` and runs enabled CLIs first (`linkedin-search`, `freehire-search`). Danish portals (`jobbank-search`, `jobdanmark-search`, `jobindex-search`, `jobnet-search`) ship with `enabled: false`.

## Search Sites

Primary:
- **linkedin.com/jobs** — covered by `linkedin-search` CLI (`--location "Bengaluru, Karnataka, India"`, `"Dubai, United Arab Emirates"`, `"London, United Kingdom"`, `"United States"`, etc.)
- **freehire.me** — covered by `freehire-search` CLI (ATS aggregator across ~50 platforms)

Secondary (company career pages & public ATS via WebSearch):
- Direct searches for target companies: Google, Apple, Meta, Microsoft, Anthropic, PayPal, Stripe, Goldman Sachs, JPMC, Razorpay, PayGlocal, American Express, Mastercard, Visa, Amazon, Adyen, Checkout.com, Worldpay, Fiserv, Wise, Airwallex, Revolut, Navi.

## Query Categories

### Priority 1: Senior Payments & Platform Product Management

#### Plain terms (portal CLIs)

```
Senior Product Manager payments
Principal Product Manager payments
Technical Product Manager payments
Senior Product Manager acquiring
```

#### WebSearch-only (`site:` fallback)

```
site:linkedin.com/jobs "Senior Product Manager" payments Bangalore OR Bengaluru
site:greenhouse.io OR site:lever.co OR site:ashbyhq.com "Senior Product Manager" payments
```

### Priority 2: Payments Infrastructure, Settlement, Risk & AI

#### Plain terms (portal CLIs)

```
Product Manager settlement
Product Manager payment gateway
AI Product Manager fintech
Senior Product Manager risk fraud
```

#### WebSearch-only (`site:` fallback)

```
site:linkedin.com/jobs "Product Manager" settlement OR acquiring OR tokenization Bangalore
site:linkedin.com/jobs "Senior Product Manager" AI payments Dubai OR London OR Singapore
```

### Priority 3: Lead / Group PM & Adjacent Fintech (Stretch / Target Companies)

#### Plain terms (portal CLIs)

```
Lead Product Manager payments
Group Product Manager fintech
Senior Product Manager fintech
```

#### WebSearch-only (`site:` fallback)

```
site:linkedin.com/jobs "Senior Product Manager" Stripe OR PayPal OR Adyen OR Razorpay OR Visa
```

## Location Filter

1. **Primary:** Bengaluru / Bangalore, India (including hybrid).
2. **Abroad (with explicit visa sponsorship / relocation only):** UAE (Dubai), UK (London), Germany, Netherlands, Ireland, Singapore, USA. Mark sponsorship status explicitly (`STATED` / `UNVERIFIED` / `NO`).

## Language Filter

- English (Primary working language). Flag any posting requiring an undeclared local language.

## Date Filter

Only include jobs posted within the last 14 days, matching `/scrape` Step 1b.3.
