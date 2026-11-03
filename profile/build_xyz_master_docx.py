import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hr(paragraph, color_hex="CCCCCC"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(pBdr)

def build_xyz_master_cv(output_path):
    doc = docx.Document()
    
    # 0.55 in margins for compact executive layout
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
    
    # 1. Header: Name
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("AVIJIT MUKHERJEE")
    r_name.font.size = Pt(18)
    r_name.font.bold = True
    r_name.font.color.rgb = RGBColor(15, 23, 42)
    
    # 2. Target Role Headline
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Senior Payments Product Leader | Acquiring Platforms | Core Settlement Infrastructure")
    r_title.font.size = Pt(10.5)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)
    
    # Contact Details
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

    def add_bullet(text, bold_prefix=None, is_sub=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.left_indent = Inches(0.25 if not is_sub else 0.4)
        p.paragraph_format.line_spacing = 1.12
        
        r_bullet = p.add_run("•  " if not is_sub else "–  ")
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

    # 3. Executive Summary (max 4 lines)
    add_section_header("Executive Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(
        "Payments product leader with 14 years of experience across acquiring platforms, payment APIs, clearing, and core settlement infrastructure, including nearly 10 years in technical product management at Visa and Cognizant. Career spans continuous delivery from Jan 2013 to Sep 2026 (13 years 9 months) across global networks and merchant processors. Generated $30M in net revenue by launching domestic-scheme processing on event-driven microservices, protected over $22M in cumulative merchant/network revenue, and reduced manual product delivery cycles by 50% using agentic AI toolchains. Certified Scrum Product Owner (CSPO) and 2-time Visa STAR Award recipient recognized for translating complex technical network standards (ISO 8583, BASE I/II, Kafka) into resilient commercial platforms."
    )
    r_sum.font.size = Pt(9.5)

    # 4. Core Skills
    add_section_header("Core Skills")
    skills = [
        ("Product Management: ", "Product Strategy & Roadmaps, Product Discovery & Problem Framing, PRDs & Business Cases, Backlog Grooming & Prioritization, User Stories & Acceptance Criteria, Experimentation & A/B Testing, Release Management, Production Incident Response (P1/P2)."),
        ("Domain & Payments: ", "Merchant Acquiring, Payment Processing, Authorization, Clearing & Settlement, Reconciliation & Accounting, Domestic & Regional Schemes (UAE Jaywan, RuPay, KSA, CUP, PayPak, MyDebit), Tokenization (COF, MIT, CIT, Network Tokens), Interchange & Qualification Analysis, EMV Certification."),
        ("Technical & Architecture: ", "ISO 8583, Visa BASE I & BASE II, TC33.A, TC05/06, IPM, CAPN, EDC, PTS, XML, JSON, REST APIs, Payload Validation & Transformation, BIN Routing & Lookup, Kafka Event-Driven Architecture, Microservices, Payments 2.0."),
        ("Analytical & Operations: ", "Funnel & Transaction Drop-Off Analysis, SLA/KPI Tracking & Reporting, Production Defect Analysis, Interchange Optimization, Root-Cause Investigation, Settlement Regeneration."),
        ("Leadership & Governance: ", "Cross-Functional Team Leadership (managed 12-member product team), Executive Stakeholder Management, Enterprise Acquirer & Merchant Collaboration, Vendor & Network Governance.")
    ]
    for cat, items in skills:
        add_bullet(items, bold_prefix=cat)

    # 5. Professional Experience
    add_section_header("Professional Experience")
    
    def add_job_header(company, role, location, dates):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        
        r_comp = p.add_run(company)
        r_comp.font.bold = True
        r_comp.font.size = Pt(10)
        r_comp.font.color.rgb = RGBColor(15, 23, 42)
        
        r_loc = p.add_run(f" | {location}")
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

    def add_achievements_summary(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.2)
        r = p.add_run("Key Highlights for Tailoring: ")
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(71, 85, 105)
        r2 = p.add_run(text)
        r2.font.size = Pt(8.5)
        r2.font.italic = True
        r2.font.color.rgb = RGBColor(100, 116, 139)

    # Visa PM
    add_job_header("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026")
    add_bullet("Generated $30M in net revenue, as measured by domestic-scheme transaction volume through CyberSource, by delivering Payments 2.0 microservices and Kafka event-driven settlement generation across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit.")
    add_bullet("Eliminated 8 FTEs of manual effort per sprint, as measured by reduced Product Development team dependencies and file turnaround under 24 hours, by automating CAS self-service testing and settlement-file retrieval for acquirers, merchants, and partners.")
    add_bullet("Reduced release-related churn and deployment delays, as measured by eliminating manual coordination across 400+ separate files per release, by transitioning settlement generation to on-demand, scoped access for acquiring partners.")
    add_bullet("Expanded regional acquiring coverage across Europe, as measured by onboarding 13 new acquirers and scaling transaction capacity by an additional 12,000 TPS, by partnering with Visa Europe on Elavon processing via VPC, simplifying JCB and Diners clearing in France and Spain via TC33.A formats, and enabling EMV processing for Greek acquirers.")
    add_bullet("Accelerated cross-border integration delivery by 67%, as measured by delivering the end-to-end integration in 2 weeks instead of the planned 6-week schedule, by using a Claude-based agentic toolchain to architect specifications, auto-generate mappings, and validate edge cases for SMCC Japan (EDQP, PIF, force-refund).")
    add_bullet("Supported the national market launch of Apple Pay in India, as measured by achieving zero processing errors and seamless settlement execution across pilot and live transactions, by conducting end-to-end certification testing and settlement validation on CyberSource while authoring product specs for Fast Refund in India and APAC.")
    add_bullet("Established the architectural blueprint for global platform modernization, as measured by defining the roadmap to decouple 85% of legacy monolithic BASE I processing into 5 scalable microservices, by authoring the Payments 2.0 Mapping Document aligned with the Visa Intelligence Authorization platform.")
    add_bullet("Standardized event-driven settlement infrastructure across distributed systems, as measured by eliminating end-of-day batch processing lags, by defining Kafka-based REST payment event schemas for Near Real-Time Feeds (NRTF), Real-Time Settlement System (RTSS), and settlement generation.")
    add_achievements_summary("$30M net revenue from multi-scheme expansion; 8 FTEs/sprint saved through CAS automation; Claude agentic toolchain cut delivery from 6 weeks to 2 weeks; onboarded 13 European acquirers (+12k TPS).")

    # Visa TPM
    add_job_header("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024")
    add_bullet("Protected $9.2M in transaction revenue, as measured by incident-linked revenue exposure prevented across CEMEA, APAC, Europe, and the Americas, by resolving high-priority settlement failures and introducing automated settlement regeneration workflows.")
    add_bullet("Grew token-based transaction volume by 25%, as measured by client processing volume across European acquirers, by streamlining implementation requirements and technical flows for Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) transactions.")
    add_bullet("Increased Visa Instalment Plan business volume by 18%, as measured by regional transaction growth in CEMEA, by simplifying technical integration specifications and partner onboarding requirements.")
    add_bullet("Enabled fuel transaction processing associated with $70M in revenue, as measured by merchant processing throughput on VPC, by mapping the Automated Fuel Dispenser (AFD) acceptance framework for US and EU markets.")
    add_bullet("Maintained release stability across 100+ acquiring partners, as measured by zero major post-release rollbacks, by leading Business Enhancement Request (BER) implementations, technical war rooms, and P1/P2 operational triage.")
    add_bullet("Reduced manual effort in product delivery by 50%, as measured by time spent authoring PRDs, business cases, and defect triage, by embedding AI-assisted workflows (Claude, Gemini) into root-cause investigations and requirement definitions.")
    add_achievements_summary("$9.2M revenue protected via settlement incident triage; 25% token volume lift in Europe; 18% instalment plan uplift in CEMEA; awarded Visa STAR Award.")

    # Visa TPA
    add_job_header("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024")
    add_bullet("Scaled financial partner onboarding to 400+ acquiring institutions, as measured by platform client growth from 100+ acquirers, by collaborating with UX and backend engineering to overhaul onboarding workflows and implement batch attribute updates.")
    add_bullet("Enabled domestic co-brand routing for Australian and Saudi Arabian markets, as measured by zero routing rejection errors in production, by defining BIN lookup algorithms and payload transformation logic for EFTPOS and MADA.")
    add_bullet("Accelerated core platform modernization, as measured by decoupling 85% of legacy monolithic BASE I validation logic across 5 dedicated microservices, by translating product business cases into Presto algorithms and authoring engineering user stories under Payments 2.0.")
    add_bullet("Ensured 100% network interchange compliance across global clearing systems, as measured by zero audit exceptions or clearing penalties, by maintaining validation and presentation formats across Visa BASE II specifications (TC33.A, TC05/06, IPM, and CAPN).")
    add_bullet("Optimized acquirer processing operations across 100+ acquiring banks, as measured by faster resolution of settlement discrepancies, by delivering technical specifications for transaction routing, settlement exceptions, and API payload transformations.")
    add_achievements_summary("Scaled platform onboarding to 400+ acquirers; decoupled 85% of BASE I validation into 5 microservices; defined co-brand routing for EFTPOS and MADA.")

    # Cognizant Team Lead
    add_job_header("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021")
    add_bullet("Protected $11M in merchant revenue, as measured by prevented transaction drop-offs and incident exposure, by introducing upstream payment-flow validation mechanisms for enterprise clients (YUM, Burger King, McDonald’s, and petroleum brands).")
    add_bullet("Increased post-release deployment accuracy from 80% to 97%, as measured by audit logs and production metrics, by standardizing API payload verification and pre-release validation protocols across enterprise merchant platforms.")
    add_bullet("Decreased P1 production incidents by 19%, as measured by incident frequency reports for high-volume enterprise merchants, by instituting early-stage upstream validation and coordinating cross-functional engineering responses.")
    add_bullet("Maintained enterprise service level agreements across 12 product managers and analysts, as measured by 100% on-time milestone delivery, by implementing structured KPI tracking dashboards, mentoring team members, and aligning delivery roadmaps with client leadership.")
    add_bullet("Accelerated production incident resolution for Fortune 500 merchants, as measured by a 40% reduction in Mean Time to Resolution (MTTR), by coordinating cross-functional engineering, processor, and merchant operations teams.")
    add_achievements_summary("$11M merchant revenue protected; post-release accuracy increased from 80% to 97%; P1 incidents reduced by 19%; people leadership of 12 PMs and analysts.")

    # Cognizant PM
    add_job_header("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019")
    add_bullet("Shortened merchant certification project cycles by 57%, as measured by time-to-market across technical onboarding tracks within five months, by architecting and launching an automated transaction-validation platform covering 13 technical specifications.")
    add_bullet("Contributed to a 32% year-end revenue increase for First Data, as measured by annual processor financial reports, by eliminating certification bottlenecks and accelerating merchant go-live dates.")
    add_bullet("Saved 10 FTEs of recurring engineering effort, as measured by reduced manual test execution hours, by automating transaction payload validation using Automation Anywhere RPA, Java, ASP.NET, and SQL.")
    add_bullet("Protected $2M+ in client revenue, as measured by identified and prevented billing discrepancies, by introducing proactive client gap validation and payload remediation workflows.")
    add_bullet("Increased transaction validation accuracy for enterprise merchants (McDonald’s, ExxonMobil, NCR), as measured by achieving a 99%+ automated validation pass rate, by establishing automated CI/CD testing pipelines and friendly-merchant pilot validation.")
    add_achievements_summary("57% reduction in certification cycle time; 32% First Data revenue increase; 10 FTEs automated effort; multiple Cognizant Annual Employee Awards.")

    # Cognizant Cert PM / SME
    add_job_header("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017")
    add_bullet("Delivered 100+ merchant and ISV Class B certifications across enterprise brands (Walmart, McDonald’s, ExxonMobil, Verifone, NCR), as measured by achieving a 97% first-pass certification success rate, by managing end-to-end credential intake, test card execution, and settlement validation.")
    add_bullet("Recognized as Cards & Payments Subject Matter Expert in October 2015, as measured by official organizational SME certification, by establishing mastery across 10+ payment specifications (Buypass, North, Nashville, ISO, EDC, and XML).")
    add_bullet("Reduced merchant clearing errors and interchange downgrades, as measured by zero non-compliance penalties on client portfolios, by analyzing mainframe settlement files, interchange qualification rules, and network presentation formats.")
    add_bullet("Accelerated EMV migration for retail and petroleum merchants, as measured by certifying 100% of target pilot locations ahead of mandate deadlines, by coordinating EMV Level 1 and Level 2 terminal certification and Automated Fuel Dispenser (AFD) test execution.")
    add_bullet("Upskilled certification engineering cohorts, as measured by successfully training 20+ junior analysts with 30% faster ramp-up time, by developing standardized training curricula covering clearing, reconciliation, network specs, and chargeback workflows.")
    add_achievements_summary("Managed 100+ enterprise certifications with 97% first-pass rate; recognized SME across 10+ payment specifications; multiple Annual Employee Awards.")

    # Cognizant PE
    add_job_header("COGNIZANT", "Process Executive / Senior Process Executive", "Kolkata, India", "Jan 2013 – Jan 2014")
    add_bullet("Maintained 100% compliance with identity governance standards, as measured by zero audit exceptions, by provisioning user roles and access permissions in the User Access Management System for Nationstar Mortgage LLC.")
    add_bullet("Accelerated user provisioning turnaround times, as measured by a 30% reduction in access request backlogs, by streamlining verification procedures and access-ticket routing.")
    add_bullet("Supported initial card transaction validation workflows for First Data, as measured by successfully completing 100% of assigned test batches, by auditing test transactions against Buypass platform specifications to prepare for merchant certification.")
    add_achievements_summary("Foundation in access governance and card transaction validation.")

    # 6. Selected Projects / Products
    add_section_header("Selected Projects / Products")
    projects = [
        ("Payments 2.0 Domestic Scheme Platform (CyberSource): ", "Architected multi-scheme processing and Kafka event-driven settlement across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit, taking platform from 0 to live production and generating $30M in net revenue."),
        ("CAS Self-Service Settlement Automation (Visa): ", "Automated settlement-file retrieval and self-service testing, moving from centralized manual generation (400+ files/release) to on-demand access in under 24 hours, saving 8 FTEs/sprint."),
        ("European Acquiring Expansion (VPC): ", "Integrated Elavon and enabled regional clearing formats across France, Spain, and Greece, onboarding 13 new acquirers and scaling platform capacity by 12,000 TPS."),
        ("Automated Fuel Dispenser (AFD) Modernization (VPC): ", "Defined US/EU AFD acceptance logic and network mapping for petroleum processing on Visa Platform Connect, supporting business associated with $70M in revenue."),
        ("Agentic AI Specification Engine (Visa / SMCC Japan): ", "Pioneered Claude-based agentic workflows to automate interface mapping, PRD authoring, and edge-case resolution, compressing delivery timeline by 67% (from 6 weeks to 2 weeks)."),
        ("Apple Pay India Certification (CyberSource): ", "Executed full market launch testing and settlement validation with 0 processing errors, authored Fast Refund specs for India and APAC."),
        ("First Data Transaction Validation Engine (Cognizant): ", "Built an algorithmic RPA and Java engine validating 13 technical specifications, accelerating certification cycle times by 57% and saving 10 FTEs.")
    ]
    for p_title, p_desc in projects:
        add_bullet(p_desc, bold_prefix=p_title)

    # 7. Education & Certifications
    add_section_header("Education & Certifications")
    p_ed = doc.add_paragraph()
    p_ed.paragraph_format.space_before = Pt(2)
    p_ed.paragraph_format.space_after = Pt(2)
    p_ed.paragraph_format.line_spacing = 1.12
    r_ed = p_ed.add_run("Bachelor of Technology (B.Tech) – Information Technology | West Bengal University of Technology (2012) | GPA: 8.6\n")
    r_ed.font.size = Pt(9.5)
    r_ed_certs = p_ed.add_run("Certifications: Certified Scrum Product Owner (CSPO) – Scrum Alliance | Product Lifecycle Management (PLM) – Visa | VPL Core Product – Visa")
    r_ed_certs.font.size = Pt(9)
    r_ed_certs.font.color.rgb = RGBColor(71, 85, 105)

    # 8. Additional Information
    add_section_header("Additional Information")
    add_bullet("2-time Visa STAR Award recipient; multiple Cognizant Annual Employee Awards; recognized Cards & Payments Subject Matter Expert (Cognizant, Oct 2015).", bold_prefix="Honors & Awards: ")
    add_bullet("English (Professional / Fluent), Bengali (Native), Hindi (Professional Working).", bold_prefix="Languages: ")
    add_bullet("Jira, Postman, Power BI, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, OpenAI, Microsoft Copilot.", bold_prefix="Technical Toolkit: ")

    doc.save(output_path)
    print(f"XYZ Master CV saved to: {output_path}")

if __name__ == "__main__":
    cv_path = r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv.docx"
    build_xyz_master_cv(cv_path)
