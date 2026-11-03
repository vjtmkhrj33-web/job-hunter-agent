import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hr(paragraph, color_hex="CCCCCC"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(pBdr)

def build_docx_from_data(title_role, summary, core_skills, exp_bullets, highlights, output_path):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.65)
        section.right_margin = Inches(0.65)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(9.5)
    font.color.rgb = RGBColor(33, 37, 41)
    
    # 1. Header
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("AVIJIT MUKHERJEE")
    r_name.font.size = Pt(18)
    r_name.font.bold = True
    r_name.font.color.rgb = RGBColor(15, 23, 42)
    
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run(title_role)
    r_title.font.size = Pt(10.5)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)
    
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(8)
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_contact = p_contact.add_run("Bengaluru, India  |  +91 9038415509  |  vjtmkhrj33@gmail.com  |  linkedin.com/in/avijitmukherjee5889/")
    r_contact.font.size = Pt(9)
    r_contact.font.color.rgb = RGBColor(71, 85, 105)
    add_hr(p_contact, "94A3B8")
    
    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(9)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title.upper())
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        add_hr(p, "CBD5E1")
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.line_spacing = 1.12
        r_bullet = p.add_run("•  ")
        r_bullet.font.size = Pt(9)
        r_bullet.font.color.rgb = RGBColor(30, 58, 138)
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.bold = True
            r_bold.font.size = Pt(9.5)
            r_bold.font.color.rgb = RGBColor(15, 23, 42)
        r_text = p.add_run(text)
        r_text.font.size = Pt(9.5)
        r_text.font.color.rgb = RGBColor(33, 37, 41)
        return p

    # Summary
    add_section_header("Executive Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(summary)
    r_sum.font.size = Pt(9.5)

    # Core Skills
    add_section_header("Core Competencies")
    for cat, items in core_skills:
        add_bullet(items, bold_prefix=cat)

    # Professional Experience
    add_section_header("Professional Experience")
    for company, role, loc, dates, bullets in exp_bullets:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        r_comp = p.add_run(company)
        r_comp.font.bold = True
        r_comp.font.size = Pt(10)
        r_comp.font.color.rgb = RGBColor(15, 23, 42)
        r_loc = p.add_run(f" | {loc}")
        r_loc.font.size = Pt(9.5)
        r_loc.font.color.rgb = RGBColor(100, 116, 139)

        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(2)
        p_sub.paragraph_format.keep_with_next = True
        r_role = p_sub.add_run(role)
        r_role.font.bold = True
        r_role.font.italic = True
        r_role.font.size = Pt(9.5)
        r_role.font.color.rgb = RGBColor(30, 58, 138)
        r_dates = p_sub.add_run(f"  ({dates})")
        r_dates.font.size = Pt(9)
        r_dates.font.italic = True
        r_dates.font.color.rgb = RGBColor(100, 116, 139)

        for b in bullets:
            add_bullet(b)

    # Education & Certs
    add_section_header("Education & Certifications")
    p_ed = doc.add_paragraph()
    p_ed.paragraph_format.space_before = Pt(2)
    p_ed.paragraph_format.space_after = Pt(2)
    p_ed.paragraph_format.line_spacing = 1.12
    r_ed = p_ed.add_run("B.Tech – Information Technology | West Bengal University of Technology (2012) | GPA: 8.6\n")
    r_ed.font.size = Pt(9.5)
    r_ed_certs = p_ed.add_run("Certifications: Certified Scrum Product Owner (CSPO) – Scrum Alliance | Product Lifecycle Management (PLM) – Visa | VPL Core Product – Visa")
    r_ed_certs.font.size = Pt(9)
    r_ed_certs.font.color.rgb = RGBColor(71, 85, 105)

    # Additional
    add_section_header("Additional Information")
    add_bullet("2-time Visa STAR Award recipient; multiple Cognizant Annual Employee Awards; recognized Cards & Payments Subject Matter Expert (Cognizant, Oct 2015).", bold_prefix="Honors & Awards: ")
    add_bullet("English (Professional / Fluent), Bengali (Native), Hindi (Professional Working).", bold_prefix="Languages: ")
    add_bullet("Jira, Postman, Power BI, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, OpenAI, Microsoft Copilot.", bold_prefix="Technical Toolkit: ")

    doc.save(output_path)
    print(f"Saved: {output_path}")

# ==============================================================================
# 1. GENERALIST PRODUCT MANAGEMENT VARIANT
# ==============================================================================
def create_generalist_variant():
    title = "Senior Product Manager | Technical Product Management & Platform Delivery"
    summary = (
        "Accomplished Senior Product Manager with nearly 10 years in product management and 14 years in enterprise software and financial platforms. "
        "Proven record of taking complex 0-to-1 products from customer discovery to commercial scale, driving $30M in net revenue, automating operations to save 8 FTEs/sprint, "
        "and cutting product delivery timelines by 67% using modern AI-assisted toolchains. Certified Scrum Product Owner (CSPO) and 2-time Visa STAR Award recipient "
        "skilled at backlog prioritization, cross-functional delivery across UX, engineering, and compliance, and scaling platforms to 400+ enterprise clients."
    )
    skills = [
        ("Product Strategy & Roadmaps: ", "0→1 Product Delivery, Vision & Strategy, Product Discovery, Customer Interviews, Backlog Grooming & Prioritization, User Stories & Acceptance Criteria, CSPO Agile Frameworks"),
        ("Execution & Analytics: ", "Release Management, Sprint Planning, A/B Testing, Funnel Optimization, KPI/SLA Governance, Incident Triage (P1/P2), Mean Time to Resolution (MTTR) Reduction"),
        ("Technical & Platform Architecture: ", "REST APIs, Event-Driven Architecture, Microservices Modernization, Monolith Decomposition, System Integration, Enterprise Scalability"),
        ("AI & Automation: ", "AI-Assisted Product Workflows (Claude, Gemini), Agentic Solution Design, Automated Requirements Generation, Robotic Process Automation (Automation Anywhere)"),
        ("Leadership & Stakeholders: ", "Cross-Functional Team Leadership (managed 12-member product team), Executive Stakeholder Communication, Partner & Vendor Collaboration")
    ]
    exp = [
        ("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026", [
            "Generated $30M in net revenue, as measured by multi-region transaction volume, by defining product requirements and delivering microservices architecture across 6 regional markets.",
            "Eliminated 8 FTEs of manual effort per sprint, as measured by reducing internal team dependencies and cutting turnaround from days to under 24 hours, by launching self-service capabilities for partners.",
            "Accelerated cross-border product delivery by 67%, as measured by delivering in 2 weeks instead of 6 weeks, by implementing an agentic AI toolchain to automate specification validation and edge-case resolution.",
            "Reduced release-related deployment delays, as measured by eliminating manual coordination across 400+ separate release files, by transitioning from batch releases to on-demand access.",
            "Supported market launch for Apple Pay India, as measured by achieving zero processing errors and seamless settlement execution across transactions, by directing end-to-end certification testing."
        ]),
        ("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024", [
            "Protected $9.2M in revenue, as measured by incident-linked exposure prevented, by resolving high-priority system bottlenecks and introducing automated data regeneration workflows.",
            "Grew recurring transaction business volume by 25%, as measured by client volume uplift, by redesigning and streamlining integration specifications for enterprise European partners.",
            "Increased regional business volume by 18%, as measured by quarterly transaction growth in CEMEA, by simplifying product onboarding requirements for financial partners.",
            "Reduced manual effort in product delivery by 50%, as measured by time spent authoring PRDs and conducting defect triage, by embedding AI-assisted workflows into weekly product discovery."
        ]),
        ("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024", [
            "Scaled enterprise partner onboarding to 400+ financial institutions, as measured by platform client growth from 100+ partners, by collaborating with UX and backend engineering on workflow automation.",
            "Accelerated core platform performance and scalability, as measured by decoupling 85% of legacy monolithic validation logic into 5 dedicated microservices, by translating business cases into engineering user stories.",
            "Ensured 100% regulatory and network compliance, as measured by zero audit exceptions or operational penalties, by maintaining international system presentation formats."
        ]),
        ("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021", [
            "Protected $11M in enterprise revenue, as measured by prevented platform drop-offs, by introducing upstream validation mechanisms before major software releases.",
            "Increased post-release deployment accuracy from 80% to 97% and decreased P1 production incidents by 19%, by standardizing API payload verification protocols across 12 product team members.",
            "Maintained 100% on-time milestone delivery across enterprise partner roadmaps, by establishing structured KPI dashboards and leading agile release ceremonies."
        ]),
        ("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019", [
            "Shortened enterprise onboarding cycle time by 57%, as measured by time-to-market within five months of launch, by architecting and launching an automated platform covering 13 technical specifications.",
            "Contributed to a 32% year-end revenue increase for First Data, as measured by annual corporate financial reports, by accelerating client go-live timelines.",
            "Saved 10 FTEs of recurring manual effort, as measured by engineering hours saved, by automating data payload validation using RPA and SQL-based rule engines."
        ]),
        ("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017", [
            "Delivered 100+ merchant and ISV certifications across enterprise brands (Walmart, McDonald’s, ExxonMobil, NCR), achieving a 97% first-pass certification success rate.",
            "Recognized as Subject Matter Expert across enterprise payment specifications; upskilled 20+ junior team members with 30% faster onboarding time."
        ]),
        ("COGNIZANT", "Process Executive / Senior Process Executive", "Kolkata, India", "Jan 2013 – Jan 2014", [
            "Maintained 100% compliance with identity governance standards and access management protocols for enterprise financial client portfolios.",
            "Reduced access provisioning backlogs by 30% by streamlining request verification and ticket resolution workflows."
        ])
    ]
    md_content = f"""# AVIJIT MUKHERJEE
**{title}**
Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjee5889/

---

## EXECUTIVE SUMMARY
{summary}

---

## CORE COMPETENCIES
{chr(10).join([f"- **{cat}** {items}" for cat, items in skills])}

---

## PROFESSIONAL EXPERIENCE
"""
    for comp, role, loc, dates, bullets in exp:
        md_content += f"\n### {comp}\n**{role}** | {loc}  \n*{dates}*\n\n"
        for b in bullets:
            md_content += f"- {b}\n"

    md_content += """
---

## EDUCATION & CERTIFICATIONS
- **Bachelor of Technology (B.Tech) – Information Technology** | West Bengal University of Technology (2012) | GPA: 8.6
- **Certified Scrum Product Owner (CSPO)** – Scrum Alliance
- **Product Lifecycle Management (PLM)** – Visa | **VPL Core Product** – Visa

---

## ADDITIONAL INFORMATION
- **Honors & Awards:** 2-time Visa STAR Award recipient; multiple Cognizant Annual Employee Awards; recognized SME (Cognizant).
- **Languages:** English (Professional / Fluent), Bengali (Native), Hindi (Professional Working).
- **Technical Toolkit:** Jira, Postman, Power BI, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, OpenAI, Microsoft Copilot.

---

### VARIANT USAGE GUIDELINES
- **Best Job Titles:** Senior Product Manager, Technical Product Manager, Platform Product Manager, Lead Product Manager.
- **Primary Keywords:** Product Strategy, Roadmap, Discovery, PRD, Backlog Prioritization, User Stories, CSPO, Microservices, Event-Driven, AI Workflows.
- **Where NOT to Use:** Highly specialized payments-only regulatory roles (use Variant 3) or Director/VP executive roles (use Variant 2).
"""
    with open(r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_generalist.md", "w", encoding="utf-8") as f:
        f.write(md_content)
    build_docx_from_data(title, summary, skills, exp, [], r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_generalist.docx")

# ==============================================================================
# 2. PRODUCT LEADERSHIP VARIANT (GPM / Head of Product / Lead PM)
# ==============================================================================
def create_leadership_variant():
    title = "Product Leader | Group Product Manager | Head of Platform Delivery"
    summary = (
        "Product Leader with 14 years of experience across platform modernization, enterprise scale, and digital product delivery, including nearly 10 years leading product teams and initiatives at Visa and Cognizant. "
        "Delivered $30M in new net revenue, protected over $22M in high-risk revenue, and managed high-performing teams of 12 PMs and analysts with 97% post-release accuracy. "
        "2-time Visa STAR Award winner recognized for aligning executive strategy with high-velocity engineering execution, scaling partner ecosystems to 400+ financial institutions, "
        "and pioneering agentic AI toolchains that reduce organizational delivery cycles by 50%."
    )
    skills = [
        ("Organizational Leadership: ", "Team Building & Mentorship (led 12-member team), Cross-Functional Alignment (Eng, Design, Legal, Ops), Strategic Roadmaps, OKR/KPI Setting & Governance, Vendor & Network Governance"),
        ("Commercial & P&L Impact: ", "Revenue Generation ($30M), Revenue Protection ($22M+), ROI Assessment, Business Case Development, Enterprise Client Relationship Management"),
        ("Platform Modernization & Scale: ", "Legacy System Modernization, Monolith-to-Microservices Decomposition, Event-Driven Architecture, Global Infrastructure Scaling (400+ Acquirers, 12k TPS Expansion)"),
        ("Process & Operational Excellence: ", "Release Management & Governance, P1/P2 Incident Command, SLA Management (97% post-release accuracy), Automated Self-Service Tooling (saved 8 FTEs/sprint)"),
        ("AI Transformation: ", "Agentic AI Integration in Product Delivery, Generative AI for Requirement Engineering, AI-Driven Incident Triage, Organizational Tooling Adoption")
    ]
    exp = [
        ("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026", [
            "Delivered $30M in net revenue expansion, as measured by multi-market transaction volume, by setting product vision and aligning distributed engineering teams to deliver Payments 2.0 microservices.",
            "Eliminated 8 FTEs of manual effort per sprint, as measured by reduced Product Development team dependencies and file turnaround under 24 hours, by driving executive sponsorship and delivery of CAS self-service automation.",
            "Accelerated complex enterprise cross-border delivery by 67%, as measured by compressing a 6-week timeline into 2 weeks, by introducing an organizational Claude-based agentic AI toolchain for SMCC Japan.",
            "Architected global platform modernization strategy, as measured by defining the roadmap to decouple 85% of legacy monolithic BASE I processing into 5 scalable microservices, by co-authoring the Payments 2.0 Mapping Document.",
            "Expanded European acquiring capacity by 12,000 TPS across 13 new financial institutions, by directing cross-regional partner alignment with Visa Europe and Elavon."
        ]),
        ("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024", [
            "Protected $9.2M in mission-critical revenue, as measured by incident-linked exposure prevented across CEMEA, APAC, Europe, and Americas, by leading technical incident command and introducing automated regeneration workflows.",
            "Led operational release governance across 100+ partner banks, as measured by maintaining zero major rollbacks, by establishing technical war rooms and cross-functional SWAT response protocols.",
            "Grew recurring transaction volume by 25% in Europe and increased Visa Instalment Plan volume by 18% in CEMEA, by driving client-focused spec simplification and regional partner enablement.",
            "Awarded Visa STAR Award for exceptional cross-functional leadership, incident management, and platform reliability."
        ]),
        ("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024", [
            "Scaled financial partner onboarding to 400+ acquiring institutions, collaborating with UX and backend engineering to overhaul partner workflows.",
            "Accelerated platform modernization by translating business cases into user stories to decouple 85% of legacy monolithic BASE I logic into 5 microservices.",
            "Ensured 100% network compliance and zero clearing penalties across international BASE II specifications."
        ]),
        ("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021", [
            "Led and mentored a team of 12 product managers and business analysts, with full accountability for KPI/SLA governance, team coaching, and enterprise client delivery reporting.",
            "Protected $11M in enterprise merchant revenue, as measured by prevented transaction drop-offs and incident exposure, by establishing early-stage quality gates in payment flows for Fortune 500 brands.",
            "Increased post-release deployment accuracy from 80% to 97% and decreased P1 production incidents by 19%, by institutionalizing rigorous API verification standards across the organization."
        ]),
        ("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019", [
            "Contributed to a 32% year-end revenue increase for First Data and saved 10 FTEs of manual effort, by sponsoring and launching an automated transaction-validation engine across 13 technical specifications.",
            "Shortened enterprise client onboarding cycle times by 57%, as measured by time-to-market across technical tracks within five months, by driving agile cross-functional delivery across engineering and QA.",
            "Protected $2M+ in client revenue, as measured by identified billing discrepancies, by introducing proactive gap remediation workflows."
        ]),
        ("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017", [
            "Directed Class B certification programs across 100+ enterprise clients (Walmart, McDonald’s, ExxonMobil), achieving a 97% first-pass certification rate.",
            "Recognized as Cards & Payments Subject Matter Expert in October 2015, training and upskilling 20+ junior analysts with 30% faster ramp-up time."
        ]),
        ("COGNIZANT", "Process Executive / Senior Process Executive", "Kolkata, India", "Jan 2013 – Jan 2014", [
            "Maintained 100% compliance with identity governance standards for enterprise mortgage and financial clients with zero audit exceptions.",
            "Streamlined access verification and resolution workflows, cutting backlog turnaround times by 30%."
        ])
    ]
    md_content = f"""# AVIJIT MUKHERJEE
**{title}**
Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjee5889/

---

## EXECUTIVE SUMMARY
{summary}

---

## CORE COMPETENCIES
{chr(10).join([f"- **{cat}** {items}" for cat, items in skills])}

---

## PROFESSIONAL EXPERIENCE
"""
    for comp, role, loc, dates, bullets in exp:
        md_content += f"\n### {comp}\n**{role}** | {loc}  \n*{dates}*\n\n"
        for b in bullets:
            md_content += f"- {b}\n"

    md_content += """
---

## EDUCATION & CERTIFICATIONS
- **Bachelor of Technology (B.Tech) – Information Technology** | West Bengal University of Technology (2012) | GPA: 8.6
- **Certified Scrum Product Owner (CSPO)** – Scrum Alliance
- **Product Lifecycle Management (PLM)** – Visa | **VPL Core Product** – Visa

---

## ADDITIONAL INFORMATION
- **Honors & Awards:** 2-time Visa STAR Award recipient; multiple Cognizant Annual Employee Awards; recognized SME (Cognizant).
- **Languages:** English (Professional / Fluent), Bengali (Native), Hindi (Professional Working).
- **Technical Toolkit:** Jira, Postman, Power BI, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, OpenAI, Microsoft Copilot.

---

### VARIANT USAGE GUIDELINES
- **Best Job Titles:** Group Product Manager (GPM), Principal Product Manager, Head of Product, Director of Product Management, Lead Staff PM.
- **Primary Keywords:** Product Leadership, Team Management, P&L Impact, Revenue Protection, Executive Governance, Scale, Modernization, Mentorship.
- **Where NOT to Use:** Junior/IC Product Analyst roles, hands-on junior coding roles.
"""
    with open(r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_leadership.md", "w", encoding="utf-8") as f:
        f.write(md_content)
    build_docx_from_data(title, summary, skills, exp, [], r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_leadership.docx")

# ==============================================================================
# 3. PAYMENTS & FINTECH DOMAIN SPECIALIST VARIANT
# ==============================================================================
def create_payments_variant():
    title = "Senior Payments Product Leader | Acquiring Platforms | Core Settlement Infrastructure"
    summary = (
        "Senior Payments Product Leader with 14 years of deep domain expertise across merchant acquiring, network specifications, transaction processing, and core settlement infrastructure. "
        "Delivered $30M in net revenue via domestic payment schemes (UAE Jaywan, RuPay, KSA), supported $70M in fuel transaction volume, and protected over $22M in merchant revenue through automated settlement regeneration. "
        "Combines deep mastery of ISO 8583, Visa BASE I/II, network tokenization, and Kafka event-driven clearing with hands-on agentic AI workflows that cut delivery schedules by 67%."
    )
    skills = [
        ("Merchant Acquiring & Processing: ", "Acquirer Onboarding (400+ banks), Authorization, Clearing & Settlement, Multi-Rail Routing, Domestic Schemes (UAE Jaywan, RuPay, KSA, CUP, PayPak), Interchange & Qualification Analysis"),
        ("Tokenization & Lifecycle: ", "Network Tokenization, Credential-on-File (COF), Merchant-Initiated (MIT), Customer-Initiated (CIT), EMV Level 1/2 Certification, BIN Lookup & Co-Brand Routing"),
        ("Payments Infrastructure & Architecture: ", "ISO 8583, Visa BASE I & BASE II, TC33.A, TC05/06, IPM, CAPN, EDC, PTS, XML, JSON, Microservices, Kafka Event-Driven Architecture, Payments 2.0"),
        ("Incident Response & Revenue Protection: ", "P1 Settlement Remediation ($9.2M protected), Automated Settlement Regeneration, Upstream Payload Validation ($11M protected at Fiserv), MTTR Reduction"),
        ("Modern Tooling & AI Delivery: ", "Claude & Gemini Agentic Toolchains, Jira, Power BI, Postman, SSMS, IntelliJ, Automation Anywhere RPA, CSPO Agile Delivery")
    ]
    exp = [
        ("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026", [
            "Generated $30M in net revenue, as measured by domestic-scheme transaction volume through CyberSource, by delivering Payments 2.0 microservices and Kafka event-driven settlement generation across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit.",
            "Eliminated 8 FTEs of manual effort per sprint, as measured by reduced Product Development team dependencies and file turnaround under 24 hours, by automating CAS self-service testing and settlement-file retrieval for acquirers, merchants, and partners.",
            "Expanded regional acquiring coverage across Europe, as measured by onboarding 13 new acquirers and scaling transaction capacity by an additional 12,000 TPS, by partnering with Visa Europe on Elavon processing via VPC, simplifying JCB and Diners clearing in France and Spain via TC33.A formats, and enabling EMV processing for Greek acquirers.",
            "Accelerated cross-border integration delivery by 67%, as measured by delivering the end-to-end integration in 2 weeks instead of the planned 6-week schedule, by using a Claude-based agentic toolchain to architect specifications, auto-generate mappings, and validate edge cases for SMCC Japan (EDQP, PIF, force-refund).",
            "Supported the national market launch of Apple Pay in India, as measured by achieving zero processing errors and seamless settlement execution across pilot and live transactions, by conducting end-to-end certification testing and settlement validation on CyberSource while authoring product specs for Fast Refund in India and APAC.",
            "Established the architectural blueprint for global platform modernization, as measured by defining the roadmap to decouple 85% of legacy monolithic BASE I processing into 5 scalable microservices, by authoring the Payments 2.0 Mapping Document aligned with the Visa Intelligence Authorization platform."
        ]),
        ("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024", [
            "Protected $9.2M in transaction revenue, as measured by incident-linked revenue exposure prevented across CEMEA, APAC, Europe, and the Americas, by resolving high-priority settlement failures and introducing automated settlement regeneration workflows.",
            "Grew token-based transaction volume by 25%, as measured by client processing volume across European acquirers, by streamlining implementation requirements and technical flows for Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) transactions.",
            "Increased Visa Instalment Plan business volume by 18%, as measured by regional transaction growth in CEMEA, by simplifying technical integration specifications and partner onboarding requirements.",
            "Enabled fuel transaction processing associated with $70M in revenue, as measured by merchant processing throughput on VPC, by mapping the Automated Fuel Dispenser (AFD) acceptance framework for US and EU markets."
        ]),
        ("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024", [
            "Scaled financial partner onboarding to 400+ acquiring institutions, as measured by platform client growth from 100+ acquirers, by collaborating with UX and backend engineering to overhaul onboarding workflows and implement batch attribute updates.",
            "Enabled domestic co-brand routing for Australian and Saudi Arabian markets, as measured by zero routing rejection errors in production, by defining BIN lookup algorithms and payload transformation logic for EFTPOS and MADA.",
            "Accelerated core platform modernization, as measured by decoupling 85% of legacy monolithic BASE I validation logic across 5 dedicated microservices, by translating product business cases into Presto algorithms and authoring engineering user stories under Payments 2.0.",
            "Ensured 100% network interchange compliance across global clearing systems, as measured by zero audit exceptions or clearing penalties, by maintaining validation and presentation formats across Visa BASE II specifications (TC33.A, TC05/06, IPM, and CAPN)."
        ]),
        ("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021", [
            "Protected $11M in merchant revenue, as measured by prevented transaction drop-offs and incident exposure, by introducing upstream payment-flow validation mechanisms for enterprise clients (YUM, Burger King, McDonald’s, and petroleum brands).",
            "Increased post-release deployment accuracy from 80% to 97% and decreased P1 production incidents by 19%, by standardizing API payload verification and pre-release validation protocols across enterprise merchant platforms.",
            "Led a 12-member team of product managers and analysts supporting enterprise financial platforms, with full accountability for KPI/SLA tracking, stakeholder updates, and release governance."
        ]),
        ("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019", [
            "Shortened merchant certification project cycles by 57%, as measured by time-to-market across technical onboarding tracks within five months, by architecting and launching an automated transaction-validation platform covering 13 technical specifications.",
            "Contributed to a 32% year-end revenue increase for First Data and saved 10 FTEs of recurring engineering effort, by automating transaction payload validation using Automation Anywhere RPA, Java, ASP.NET, and SQL.",
            "Protected $2M+ in client revenue, as measured by identified and prevented billing discrepancies, by introducing proactive client gap validation and payload remediation workflows."
        ]),
        ("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017", [
            "Delivered 100+ merchant and ISV Class B certifications across enterprise brands (Walmart, McDonald’s, ExxonMobil, Verifone, NCR), as measured by achieving a 97% first-pass certification success rate, by managing end-to-end credential intake, test card execution, and settlement validation.",
            "Recognized as Cards & Payments Subject Matter Expert in October 2015, establishing mastery across 10+ payment specifications (Buypass, North, Nashville, ISO, EDC, and XML)."
        ]),
        ("COGNIZANT", "Process Executive / Senior Process Executive", "Kolkata, India", "Jan 2013 – Jan 2014", [
            "Maintained 100% compliance with identity governance standards, as measured by zero audit exceptions, in user access management for mortgage and financial clients.",
            "Accelerated access provisioning turnaround by 30% through streamlined verification procedures and ticket routing.",
            "Supported card transaction validation for First Data by auditing test transactions against Buypass platform specifications to prepare for merchant certification."
        ])
    ]
    md_content = f"""# AVIJIT MUKHERJEE
**{title}**
Bengaluru, India | +91 9038415509 | vjtmkhrj33@gmail.com | linkedin.com/in/avijitmukherjee5889/

---

## EXECUTIVE SUMMARY
{summary}

---

## CORE COMPETENCIES
{chr(10).join([f"- **{cat}** {items}" for cat, items in skills])}

---

## PROFESSIONAL EXPERIENCE
"""
    for comp, role, loc, dates, bullets in exp:
        md_content += f"\n### {comp}\n**{role}** | {loc}  \n*{dates}*\n\n"
        for b in bullets:
            md_content += f"- {b}\n"

    md_content += """
---

## EDUCATION & CERTIFICATIONS
- **Bachelor of Technology (B.Tech) – Information Technology** | West Bengal University of Technology (2012) | GPA: 8.6
- **Certified Scrum Product Owner (CSPO)** – Scrum Alliance
- **Product Lifecycle Management (PLM)** – Visa | **VPL Core Product** – Visa

---

## ADDITIONAL INFORMATION
- **Honors & Awards:** 2-time Visa STAR Award recipient; multiple Cognizant Annual Employee Awards; recognized SME (Cognizant).
- **Languages:** English (Professional / Fluent), Bengali (Native), Hindi (Professional Working).
- **Technical Toolkit:** Jira, Postman, Power BI, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, OpenAI, Microsoft Copilot.

---

### VARIANT USAGE GUIDELINES
- **Best Job Titles:** Senior Payments Product Manager, Core Settlement PM, Acquiring Platform PM, Head of Payments Infrastructure.
- **Primary Keywords:** Acquiring, Merchant Processing, Settlement & Clearing, ISO 8583, BASE I/II, Tokenization, Domestic Schemes, UAE Jaywan, Kafka.
- **Where NOT to Use:** Non-fintech consumer apps, ad-tech, pure marketing PM roles.
"""
    with open(r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_payments_ai.md", "w", encoding="utf-8") as f:
        f.write(md_content)
    build_docx_from_data(title, summary, skills, exp, [], r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv_payments_ai.docx")

if __name__ == "__main__":
    create_generalist_variant()
    create_leadership_variant()
    create_payments_variant()

