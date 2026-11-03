import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def add_hr(paragraph, color_hex="CCCCCC"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(pBdr)

def build_master_cv(output_path):
    doc = docx.Document()
    
    # Page Setup - 0.6 in margins
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
    font.size = Pt(10)
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
    
    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(3)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Senior Payments Product Leader | Acquiring Platforms | Core Settlement Infrastructure")
    r_title.font.size = Pt(10.5)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138) # Classic Navy
    
    # Contact
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(8)
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_contact = p_contact.add_run("Bengaluru, India  |  +91 9038415509  |  vjtmkhrj33@gmail.com  |  linkedin.com/in/avijitmukherjee5889/")
    r_contact.font.size = Pt(9.5)
    r_contact.font.color.rgb = RGBColor(71, 85, 105)
    add_hr(p_contact, "94A3B8")
    
    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(title.upper())
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = RGBColor(15, 23, 42)
        add_hr(p, "CBD5E1")
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(2.5)
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.line_spacing = 1.12
        
        # bullet symbol
        r_bullet = p.add_run("•  ")
        r_bullet.font.size = Pt(9.5)
        r_bullet.font.color.rgb = RGBColor(30, 58, 138)
        
        if bold_prefix:
            r_bold = p.add_run(bold_prefix)
            r_bold.font.bold = True
            r_bold.font.size = Pt(10)
            r_bold.font.color.rgb = RGBColor(15, 23, 42)
            
        r_text = p.add_run(text)
        r_text.font.size = Pt(10)
        r_text.font.color.rgb = RGBColor(33, 37, 41)
        return p

    # --- PROFESSIONAL SUMMARY ---
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(3)
    p_sum.paragraph_format.space_after = Pt(6)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(
        "Payments product professional with 14 years of experience across acquiring platforms, merchant processing, settlement infrastructure, and payment APIs, including nearly 10 years in technical product management at Visa (CyberSource / Visa Platform Connect) and Cognizant (Fiserv / First Data). Proven track record of taking complex payment products from specification to production, contributing to $30M in net revenue through domestic-scheme implementations and supporting $70M in fuel transaction processing. Combines hands-on domain depth in ISO 8583, clearing, network tokenization, and Kafka event-driven settlement with modern AI-assisted product workflows that cut delivery cycles in half."
    )
    r_sum.font.size = Pt(10)

    # --- CORE COMPETENCIES ---
    add_section_header("Core Competencies")
    competencies = [
        ("Product Strategy & Delivery: ", "End-to-End Product Management, 0→1 Feature Delivery, Product Discovery, PRD & User Story Authoring, Acceptance Criteria, A/B Testing, Release Management, P1/P2 Incident Response"),
        ("Payments & Acquiring: ", "Merchant Acquiring, Payment Processing, Authorization, Clearing & Settlement, Reconciliation, Domestic Schemes, Acquirer Onboarding, Interchange & Qualification"),
        ("Technical Architecture: ", "ISO 8583, Visa BASE I & BASE II, TC33.A, TC05/06, IPM, API Payload Validation & Transformation, Microservices, Kafka Event-Driven Architecture, Payments 2.0"),
        ("Tokenization & Compliance: ", "Network Tokenization, Credential-on-File (COF), Merchant-Initiated Transactions (MIT), Customer-Initiated Transactions (CIT), EMV Certification, BIN Routing"),
        ("Efficiency & Tooling: ", "AI-Assisted Product Workflows, Agentic AI Solution Design, Workflow Automation, Process Optimization")
    ]
    for cat, items in competencies:
        add_bullet(items, bold_prefix=cat)

    # --- PROFESSIONAL EXPERIENCE ---
    add_section_header("Professional Experience")
    
    def add_job_header(company, role, location, dates):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.keep_with_next = True
        
        r_comp = p.add_run(company)
        r_comp.font.bold = True
        r_comp.font.size = Pt(10.5)
        r_comp.font.color.rgb = RGBColor(15, 23, 42)
        
        r_loc = p.add_run(f" | {location}")
        r_loc.font.size = Pt(10)
        r_loc.font.color.rgb = RGBColor(100, 116, 139)
        
        p_sub = doc.add_paragraph()
        p_sub.paragraph_format.space_before = Pt(0)
        p_sub.paragraph_format.space_after = Pt(3)
        p_sub.paragraph_format.keep_with_next = True
        
        r_role = p_sub.add_run(role)
        r_role.font.bold = True
        r_role.font.italic = True
        r_role.font.size = Pt(10)
        r_role.font.color.rgb = RGBColor(30, 58, 138)
        
        r_dates = p_sub.add_run(f"  ({dates})")
        r_dates.font.size = Pt(9.5)
        r_dates.font.italic = True
        r_dates.font.color.rgb = RGBColor(100, 116, 139)

    # Visa PM
    add_job_header("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026")
    add_bullet("Delivered domestic-scheme processing across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit on CyberSource using Payments 2.0 microservices and Kafka event-driven settlement generation, contributing $30M in net revenue.")
    add_bullet("Automated CAS self-service testing and settlement-file retrieval for acquirers, merchants, and technology partners, reducing file turnaround to under 24 hours and saving 8 FTEs of manual effort per sprint.")
    add_bullet("Transitioned settlement-file generation from a centralized 400+ file per Business Enhancement Request (BER) bottleneck to on-demand, scoped acquirer access, eliminating release churn.")
    add_bullet("Partnered with Visa Europe on Elavon transaction processing via VPC and Visa Intelligence Auth; streamlined JCB and Diners clearing in France and Spain using TC33.A formats.")
    add_bullet("Collaborated with SMCC Japan to architect and ship EDQP, PIF, and force-refund requirements on an accelerated delivery timeline using a Claude agentic toolchain.")
    add_bullet("Supported the India launch of Apple Pay through certification testing and settlement validation on CyberSource, and authored product specifications for Fast Refund in India and APAC.")
    add_bullet("Co-authored the Payments 2.0 Mapping Document to transition legacy processing engines onto the Visa Intelligence Authorization platform roadmap.")
    add_bullet("Defined Kafka-based REST event contracts with the Framework and Orchestration teams for Near Real-Time Feeds (NRTF), Real-Time Settlement System (RTSS), and settlement generation.")

    # Visa TPM
    add_job_header("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024")
    add_bullet("Resolved critical production settlement incidents across CEMEA, APAC, Europe, and the Americas, introducing automated regeneration workflows that contributed to $9.2M in revenue protection.")
    add_bullet("Streamlined implementation of Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) transaction flows for European clients, driving a 25% increase in token-based transaction volume.")
    add_bullet("Simplified technical requirements for the Visa Instalment Plan in CEMEA, increasing regional transaction business volume by 18%.")
    add_bullet("Mapped automated fuel dispenser (AFD) frameworks for US and EU markets, enabling fuel transaction acceptance on VPC associated with $70M in transaction revenue.")
    add_bullet("Led Business Enhancement Request (BER) operational rollouts across 100+ acquirers, including customer release notes, technical war rooms, and rollback protocols.")
    add_bullet("Integrated AI-driven log and defect analysis into daily triage, accelerating incident root-cause investigations.")

    # Visa TPA
    add_job_header("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024")
    add_bullet("Supported core platform operations for Visa Platform Connect (VPC), serving 100+ acquiring banks across transaction processing, settlement, and clearing.")
    add_bullet("Designed validation and transformation logic for EFTPOS co-brand processing and contributed to MADA co-brand transaction routing algorithms using BIN lookup mechanisms.")
    add_bullet("Translated product requirements into Presto algorithms and engineering stories to decompose monolithic BASE I processing into microservices under Payments 2.0.")
    add_bullet("Partnered with UX and engineering teams to overhaul acquirer onboarding flows with batch attribute updates, supporting platform expansion to over 400 acquirers.")
    add_bullet("Maintained network interchange formatting compliance across Visa BASE II specifications, including TC33.A, TC05/06, IPM, and CAPN standards.")

    # Cognizant Team Lead
    add_job_header("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021")
    add_bullet("Led a team of 12 product professionals supporting enterprise merchants on Fiserv processing platforms, managing SLA/KPI performance and client delivery reporting.")
    add_bullet("Introduced early-stage payment-flow validation mechanisms for high-volume enterprise merchants (YUM, Burger King, McDonald’s, and petroleum clients), preventing $11M in revenue loss.")
    add_bullet("Managed API payload verification, release cycles, and production escalations across merchants and core processor teams.")

    # Cognizant PM
    add_job_header("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019")
    add_bullet("Launched an automated transaction-validation platform translating 13 technical specifications into automated verification logic using Automation Anywhere RPA, Java, ASP.NET, and SQL.")
    add_bullet("Shortened the certification project-management cycle by 57% within five months of launch, contributing to a reported 32% year-end revenue increase for First Data and saving 10 FTEs of manual effort.")
    add_bullet("Applied preventive validation and client gap remediation that protected $2M+ in revenue for enterprise merchant clients.")

    # Cognizant Cert PM / SME
    add_job_header("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017")
    add_bullet("Managed end-to-end Class B certifications and onboarding for 100+ enterprise clients and ISVs, including McDonald’s, YUM, ExxonMobil, NCR, Verifone, and Walmart across retail, e-commerce, and petroleum sectors.")
    add_bullet("Recognized as Payments Subject Matter Expert in October 2015, supporting Buypass, North, Nashville, ISO, EDC, and XML platform specifications.")
    add_bullet("Validated clearing, settlement, interchange tariffs, bill-back codes, and chargeback rules across Visa, Mastercard, Amex, Discover, Diners, JCB, and fleet networks.")

    # Cognizant PE
    add_job_header("COGNIZANT", "Process Executive / Senior Process Executive", "Kolkata, India", "Jan 2013 – Jan 2014")
    add_bullet("Provisioned user roles and access permissions in the User Access Management System for Nationstar Mortgage LLC before transitioning to First Data cards and payments certification in early 2014.")

    # --- TECHNICAL SKILLS ---
    add_section_header("Technical Skills")
    tech_skills = [
        ("Payments Protocols & Formats: ", "ISO 8583, Visa BASE I & BASE II, TC33.A, TC05/06, IPM, CAPN, EDC, PTS, XML, JSON, EMV Level 1 & 2, Network Tokenization, BIN Lookup, Co-Brand Routing"),
        ("Architecture & Infrastructure: ", "Kafka Event-Driven Architecture, Microservices, REST APIs, Payments 2.0, Settlement & Clearing Engines, Cloud Platforms"),
        ("Tools & Platforms: ", "Jira, Power BI, Postman, SQL Server Management Studio (SSMS), IntelliJ, Automation Anywhere, Claude, Gemini, Microsoft Copilot")
    ]
    for cat, items in tech_skills:
        add_bullet(items, bold_prefix=cat)

    # --- EDUCATION ---
    add_section_header("Education")
    p_edu = doc.add_paragraph()
    p_edu.paragraph_format.space_before = Pt(3)
    p_edu.paragraph_format.space_after = Pt(4)
    r_ed1 = p_edu.add_run("Bachelor of Technology (B.Tech) – Information Technology\n")
    r_ed1.font.bold = True
    r_ed1.font.size = Pt(10)
    r_ed2 = p_edu.add_run("West Bengal University of Technology | Graduated: June 2012 | GPA: 8.6")
    r_ed2.font.size = Pt(9.5)
    r_ed2.font.color.rgb = RGBColor(71, 85, 105)

    # --- CERTIFICATIONS ---
    add_section_header("Certifications")
    certs = [
        "Certified Scrum Product Owner (CSPO) – Scrum Alliance",
        "Product Lifecycle Management (PLM) – Visa",
        "VPL Core Product – Visa"
    ]
    for c in certs:
        add_bullet(c)

    doc.save(output_path)
    print(f"Master CV saved to: {output_path}")

def build_cover_letter_doc(output_path):
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(33, 37, 41)
    
    # Header: Sender
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    r_name = p_name.add_run("AVIJIT MUKHERJEE")
    r_name.font.size = Pt(16)
    r_name.font.bold = True
    r_name.font.color.rgb = RGBColor(15, 23, 42)
    
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(12)
    r_contact = p_contact.add_run("Bengaluru, India  |  +91 9038415509  |  vjtmkhrj33@gmail.com  |  linkedin.com/in/avijitmukherjee5889/")
    r_contact.font.size = Pt(9.5)
    r_contact.font.color.rgb = RGBColor(71, 85, 105)
    add_hr(p_contact, "94A3B8")
    
    # Date & Recipient
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(6)
    p_meta.paragraph_format.space_after = Pt(10)
    p_meta.paragraph_format.line_spacing = 1.15
    r_meta = p_meta.add_run("[Date]\n\nHiring Team / Product Leadership\n[Target Company Name]\n[Target Company Office Location]")
    r_meta.font.size = Pt(10)
    r_meta.font.color.rgb = RGBColor(71, 85, 105)
    
    # Subject Line
    p_subj = doc.add_paragraph()
    p_subj.paragraph_format.space_before = Pt(4)
    p_subj.paragraph_format.space_after = Pt(10)
    r_sub_label = p_subj.add_run("Subject: ")
    r_sub_label.font.bold = True
    r_sub = p_subj.add_run("Application for Senior / Lead Product Manager – [Target Role] – Avijit Mukherjee")
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor(30, 58, 138)
    
    # Salutation
    p_sal = doc.add_paragraph()
    p_sal.paragraph_format.space_before = Pt(0)
    p_sal.paragraph_format.space_after = Pt(8)
    p_sal.add_run("Dear [Hiring Manager Name or \"Product Team\"],")
    
    # Body Paragraph 1
    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.space_before = Pt(0)
    p_b1.paragraph_format.space_after = Pt(8)
    p_b1.paragraph_format.line_spacing = 1.18
    p_b1.add_run(
        "I’ve spent the past 14 years in payment infrastructure, acquiring platforms, and settlement systems—most recently at Visa (CyberSource / Visa Platform Connect), where I led platform modernization and multi-rail payment scheme expansion. When I saw [Company Name]’s focus on [specific initiative from the JD: e.g., expanding multi-rail acquiring / scaling merchant platform APIs / modernizing settlement], this role immediately stood out because it directly overlaps with systems I have taken from technical spec to production."
    )
    
    # Body Paragraph 2 (Bridge to bullets)
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.space_before = Pt(0)
    p_b2.paragraph_format.space_after = Pt(6)
    p_b2.paragraph_format.line_spacing = 1.18
    p_b2.add_run(
        "Over my career across Visa, Fiserv, and First Data, I have focused on solving the hard plumbing of payments—APIs, ISO 8583 specs, Kafka event streams, and reconciliation—while tying engineering work directly to commercial outcomes:"
    )
    
    def add_cl_bullet(bold_label, text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.line_spacing = 1.15
        
        r_b = p.add_run("•  ")
        r_b.font.size = Pt(10)
        r_b.font.color.rgb = RGBColor(30, 58, 138)
        
        r_lbl = p.add_run(bold_label)
        r_lbl.font.bold = True
        r_lbl.font.size = Pt(10.5)
        r_lbl.font.color.rgb = RGBColor(15, 23, 42)
        
        r_t = p.add_run(text)
        r_t.font.size = Pt(10.5)
        return p

    add_cl_bullet("Domestic Scheme & Platform Modernization: ", "Delivered domestic-scheme processing across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit via CyberSource using Payments 2.0 and Kafka-driven settlement, contributing $30M in net revenue.")
    add_cl_bullet("Tokenization & Recurring Payments: ", "Boosted token-based transaction volume by 25% across European acquirers by eliminating integration friction for Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) transaction flows.")
    add_cl_bullet("Platform Automation & Operations: ", "Automated CAS self-service testing and settlement-file retrieval for acquirers and merchants, cutting delivery times from days to under 24 hours and saving 8 FTEs of manual effort per sprint.")

    # Body Paragraph 3
    p_b3 = doc.add_paragraph()
    p_b3.paragraph_format.space_before = Pt(6)
    p_b3.paragraph_format.space_after = Pt(8)
    p_b3.paragraph_format.line_spacing = 1.18
    p_b3.add_run(
        "My approach to product management is hands-on and technical: I am as comfortable reviewing ISO payloads, Kafka event definitions, and exception queues with engineering as I am discussing interchange qualification and revenue levers with business partners. I also actively integrate AI-assisted workflows into my product cycle, cutting PRD and business-case drafting time in half to keep teams shipping faster."
    )
    
    # Body Paragraph 4 (Close)
    p_b4 = doc.add_paragraph()
    p_b4.paragraph_format.space_before = Pt(0)
    p_b4.paragraph_format.space_after = Pt(14)
    p_b4.paragraph_format.line_spacing = 1.18
    p_b4.add_run(
        "I’d welcome 15 minutes to learn more about [Company Name]’s upcoming payment priorities and share how my background in core acquiring platforms can help your team execute."
    )
    
    # Sign-off
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(0)
    p_sign.paragraph_format.space_after = Pt(2)
    p_sign.add_run("Sincerely,")
    
    p_sig_name = doc.add_paragraph()
    p_sig_name.paragraph_format.space_before = Pt(8)
    p_sig_name.paragraph_format.space_after = Pt(0)
    r_sn = p_sig_name.add_run("Avijit Mukherjee")
    r_sn.font.bold = True
    
    # Page Break for Switchboard Appendix
    doc.add_page_break()
    
    p_app_title = doc.add_paragraph()
    p_app_title.paragraph_format.space_before = Pt(0)
    p_app_title.paragraph_format.space_after = Pt(4)
    r_apt = p_app_title.add_run("APPENDIX: THE 3-MINUTE PROOF BULLET SWITCHBOARD")
    r_apt.font.size = Pt(13)
    r_apt.font.bold = True
    r_apt.font.color.rgb = RGBColor(15, 23, 42)
    add_hr(p_app_title, "94A3B8")
    
    p_app_desc = doc.add_paragraph()
    p_app_desc.paragraph_format.space_before = Pt(4)
    p_app_desc.paragraph_format.space_after = Pt(8)
    r_apd = p_app_desc.add_run("Use these pre-verified, quantified bullet points to swap into the 3 bullet slots in the cover letter depending on what each employer values most:")
    r_apd.font.size = Pt(10)
    r_apd.font.italic = True
    r_apd.font.color.rgb = RGBColor(100, 116, 139)
    
    switchboard_categories = [
        ("0→1 Builds, Scheme Expansion & Architecture", [
            ("Domestic Scheme & Platform Modernization: ", "Delivered domestic-scheme processing across KSA, RuPay, CUP, Jaywan, PayPak, and MyDebit via CyberSource using Payments 2.0 and Kafka-driven settlement, contributing $30M in net revenue."),
            ("Modernization Roadmap: ", "Co-authored the Payments 2.0 mapping specifications bridging legacy monolithic architectures to event-driven services and the Visa Intelligence Authorization platform roadmap.")
        ]),
        ("Merchant Scale, Tokenization & Conversions", [
            ("Tokenization & Recurring Payments: ", "Boosted token-based transaction volume by 25% across European acquirers by eliminating integration friction for Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) transaction flows."),
            ("Regional Product Expansion: ", "Simplified implementation requirements for the Visa Instalment Plan in CEMEA, increasing regional transaction business volume by 18%.")
        ]),
        ("Platform Automation, Developer Tooling & Operations", [
            ("Self-Service & Developer Operations: ", "Automated CAS self-service testing and settlement-file retrieval for acquirers and merchants, cutting delivery times from days to under 24 hours and saving 8 FTEs of manual effort per sprint."),
            ("Certification Automation: ", "Shortened merchant certification project cycles by 57% at First Data by building an automated transaction-validation engine across 13 technical specifications, saving 10 FTEs.")
        ]),
        ("Enterprise Scale, High Volume & Specialized Processing", [
            ("High-Volume Acceptance: ", "Mapped the automated fuel dispenser (AFD) acceptance framework across US and EU markets on Visa Platform Connect, supporting transaction processing associated with $70M in revenue.")
        ]),
        ("Settlement, Reconciliation, Risk & Revenue Protection", [
            ("Settlement & Incident Remediation: ", "Protected $9.2M in revenue at Visa and $11M across enterprise merchants at Fiserv by instituting upstream payload validation, automated regeneration workflows, and rapid root-cause triage for high-priority settlement failures."),
            ("Client Gap Remediation: ", "Prevented $2M+ in revenue loss for Cognizant clients through preventive transaction validation and partner gap remediation.")
        ])
    ]
    
    for category_name, bullets in switchboard_categories:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(8)
        p_c.paragraph_format.space_after = Pt(2)
        p_c.paragraph_format.keep_with_next = True
        r_c = p_c.add_run(category_name)
        r_c.font.bold = True
        r_c.font.size = Pt(11)
        r_c.font.color.rgb = RGBColor(30, 58, 138)
        
        for lbl, txt in bullets:
            add_cl_bullet(lbl, txt)

    doc.save(output_path)
    print(f"Cover Letter Template saved to: {output_path}")

if __name__ == "__main__":
    cv_path = r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\master_cv.docx"
    cl_path = r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\profile\cover_letter_template.docx"
    build_master_cv(cv_path)
    build_cover_letter_doc(cl_path)

