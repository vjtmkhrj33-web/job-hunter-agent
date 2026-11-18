# Achievements for STAR stories (Source of Truth — Synced 2026-11-18)

For each story: Situation, Task, Action (what I personally did), Result (numbers), What I would do differently.

## Verified Master Metrics & Achievements

1. **Domestic-Scheme Expansion on CyberSource / VPC (Payments 2.0 / Kafka):**
   - Owned requirements, field mapping, and coordination for six domestic schemes: KSA, RuPay, CUP, UAE Jaywan, PayPak, and MyDebit.
   - Result: Contributed to $30M in net revenue.
2. **Agentic AI, 0 to 1 (Claude Transaction-Validation System):**
   - Led a Claude-based transaction-validation system with human review and shadow-mode evaluation frameworks protecting deterministic controls.
   - Result: Prevented $7.5M in high-risk exposure in one week of live data; p95 end-to-end latency under 1.2s on ~300-token prompts; p95 time to first token (TTFT) under 350ms.
3. **Self-Service Settlement Platform (Visa):**
   - Owned the roadmap for a self-service settlement platform serving 400+ acquiring banks processing batches up to 65K TPS within a 24-hour SLA.
   - Result: Automated certification and settlement-file delivery to under 24 hours, removing a manual dependency that regularly tied up 8 FTEs (15 at peak).
4. **AI-Assisted Delivery Acceleration (Visa):**
   - Used Claude workflows with human review for spike analysis, mapping, and regression design.
   - Result: Cut delivery schedules by about 50% on average; SMCC Japan integration cut from 6 weeks to 2 weeks (67%); per-feature staffing cut from 1 developer + 2 testers to 1 + 1 (~10 FTEs tracked in Jira).
5. **Fast Refund (India & APAC):**
   - Architected the design with the senior engineering lead and architect, led it through release, and trimmed scope to ship sooner.
   - Result: Release performance testing and monitoring showed refunds within 24 hours at a 97% success rate (clients went live after departure).
6. **Compliance Pivot (UDAAP & Fair Lending):**
   - Resolved UDAAP and fair-lending concerns by pivoting from a fully automated rule engine to a hybrid MVP.
   - Result: Zero regulatory escalations.
7. **Apple Pay India Settlement Readiness (RBI):**
   - Moved globally deployed code to the India data center (RBI) for settlement readiness.
   - Result: Zero processing errors in pilot transactions (go-live followed after departure).
8. **European Acquiring Capacity Expansion (VPC):**
   - Simplified card-scheme mapping from authorisation to settlement (Elavon processing, JCB & Diners clearing in France/Spain, EMV for Greek acquirers).
   - Result: Helped add 13 acquirers and 12,000 TPS of capacity.
9. **Microservices Modernization Roadmap (Payments 2.0):**
   - Defined the roadmap to split the legacy validation and transformation engine into five microservices (~85% of the monolith's logic), validating all API fields in one pass.
   - Result: 85% monolithic logic decoupled into 5 microservices.
10. **P1 Settlement Revenue Protection & Regeneration (Visa TPM):**
    - Resolved P1 settlement failures across CEMEA, APAC, Europe, and the Americas and added automated settlement regeneration.
    - Result: Protected $9.2M in transaction revenue exposure (attributed to Customer Support post-release).
11. **Tokenization & Visa Instalment Plan Growth (Visa TPM):**
    - Streamlined COF, MIT, and CIT flows starting from a COF logic gap not on the roadmap; simplified Visa Instalment Plan rules and launched in CEMEA first, then globally.
    - Result: Grew European token-based transaction volume by 25%; increased Visa Instalment Plan volume by 18% in CEMEA.
12. **Automated Fuel Dispenser (AFD) Acceptance Framework (Visa TPM):**
    - Mapped acceptance framework for US/EU markets and decoded Visa/Mastercard fleet specs into ISO field and clearing-file mappings across Fiserv, gateways, and PayFacs.
    - Result: Key contributor to AFD processing associated with $70M in annual revenue.
13. **Acquirer Onboarding & Domestic Co-Brand Routing (Visa TPA):**
    - Set up domestic co-brand routing for Australia (EFTPOS) and Saudi Arabia (MADA) with zero routing rejections; root-caused merchant API payload errors (+18% qualification rates, lower CEMEA scheme penalties); stepped in for absent PM on Acquirer Onboarding Platform (400+ acquirers) with architect approval.
    - Result: Accelerated partner onboarding by 25%; 100% BASE II compliance with zero audit exceptions.
14. **Team Leadership & Merchant Revenue Protection (Cognizant Team Lead):**
    - Led 12-member team of PMs and analysts (10 Kolkata, 2 Chennai; direct reports); standardized API payload verification and pre-release checks.
    - Result: Raised post-release accuracy from 80% to 97%, cut P1 incidents by 19%, delivered 97% on time, lifted monthly certifications from max 10 to 15–18 (+50% to +80%), protected an estimated $11M in merchant revenue.
15. **Automated Transaction Validator (Cognizant PM - Automation):**
    - Built 8-bot validator (Automation Anywhere, Java, ASP.NET, SQL) checking ISO host payload, EMV terminal log, and PTS batch record across Nashville, North, Buypass, and PTS.
    - Result: Cut certification cycles by 57% within 5 months; replaced 300 analyst-hours/day (~30 analysts) with a 1-hour run (25 analysts moved to Project Lead roles); 100% spec coverage under CI/CD; protected an estimated $2M+ in client revenue; Director Special award (Dec 2017).
16. **EMV Quick Chip & $1.5M Recovery (Cognizant Project Lead):**
    - Ran certification for 100+ merchants/ISVs at 97% first-pass rate coordinating a 30-member analyst team; raised quick-service chain's Quick Chip approval rate from 68% to 95% in 6 weeks (15K+ txns/day); traced auth-to-settlement code gap to recover $1.5M in ~1 month (2016).
    - Result: SME Award (2016) for validating 1,000+ transactions in a single day; Cards & Payments SME (Oct 2015).
17. **Awards & Recognition:**
    - Visa STAR Award 2024 and 2025 (KSA launch success; automating certification through self-service solution).
    - Cognizant Annual Awards: Best Employee 2013–2018 (six consecutive years).
    - Director Special (Dec 2017); SME Award (2016); 5th at University (B.Tech, GPA 8.6/10).
