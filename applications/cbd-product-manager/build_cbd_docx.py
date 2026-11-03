import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def add_hr(paragraph, color_hex="CCCCCC"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="2" w:color="{color_hex}"/></w:pBdr>')
    pPr.append(pBdr)

def build_cbd_tailored_cv(output_path):
    doc = docx.Document()
    
    # 0.55 in margins for compact 1-page fit
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Calibri'
    font.size = Pt(9.5)
    font.color.rgb = RGBColor(33, 37, 41)
    
    # Header: Name
    p_name = doc.add_paragraph()
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(2)
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_name = p_name.add_run("AVIJIT MUKHERJEE")
    r_name.font.size = Pt(17)
    r_name.font.bold = True
    r_name.font.color.rgb = RGBColor(15, 23, 42)
    
    # Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("Senior Product Manager – Digital Banking, Payments & AI Platform Delivery")
    r_title.font.size = Pt(10.5)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(30, 58, 138)
    
    # Contact
    p_contact = doc.add_paragraph()
    p_contact.paragraph_format.space_before = Pt(0)
    p_contact.paragraph_format.space_after = Pt(6)
    p_contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_contact = p_contact.add_run("Bengaluru, India  |  +91 9038415509  |  vjtmkhrj33@gmail.com  |  linkedin.com/in/avijitmukherjee5889/")
    r_contact.font.size = Pt(9)
    r_contact.font.color.rgb = RGBColor(71, 85, 105)
    add_hr(p_contact, "94A3B8")
    
    def add_section_header(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
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
        p.paragraph_format.left_indent = Inches(0.2)
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

    # --- SUMMARY ---
    add_section_header("Professional Summary")
    p_sum = doc.add_paragraph()
    p_sum.paragraph_format.space_before = Pt(2)
    p_sum.paragraph_format.space_after = Pt(4)
    p_sum.paragraph_format.line_spacing = 1.15
    r_sum = p_sum.add_run(
        "Payments and digital banking product professional with 14 years of experience across acquiring platforms, transaction processing, payment APIs, and settlement infrastructure, including nearly 10 years in product management at Visa (CyberSource / Visa Platform Connect) and Cognizant (Fiserv / First Data). Proven track record delivering regional payment schemes across CEMEA—including the UAE’s domestic scheme (Jaywan)—generating $30M in net revenue, and embedding agentic AI workflows to reduce product discovery and PRD cycles by 50%. Certified Scrum Product Owner (CSPO) with deep experience leading agile sprint teams, managing backlogs, and aligning cross-functional stakeholders across IT, compliance, and business operations."
    )
    r_sum.font.size = Pt(9.5)

    # --- CORE COMPETENCIES ---
    add_section_header("Core Competencies")
    competencies = [
        ("Digital Banking & Product Strategy: ", "Digital Banking Channels, Mobile & Online Payments, Product Vision & Roadmaps, Backlog Grooming, User Stories & Acceptance Criteria, Value Creation & ROI Assessment"),
        ("Payments & Regional Schemes: ", "UAE Domestic Scheme (Jaywan), CEMEA Card Schemes, Acquiring & Merchant Processing, Settlement & Reconciliation, Payment APIs & Partner Integrations, ISO 8583, BASE I/II"),
        ("AI & Modern Platform Delivery: ", "Enterprise AI Adoption, Agentic AI Solution Design, Automated Requirements Generation, Kafka Event-Driven Architecture, Microservices, Platform Modernization"),
        ("Agile Leadership & Governance: ", "Certified Scrum Product Owner (CSPO), Cross-Functional Team Leadership, Jira Backlog Management, Release Planning & Coordination, QA & UX Alignment, P1/P2 Incident Management")
    ]
    for cat, items in competencies:
        add_bullet(items, bold_prefix=cat)

    # --- PROFESSIONAL EXPERIENCE ---
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

    # Visa PM
    add_job_header("VISA", "Product Manager – CyberSource / Visa Platform Connect (VPC)", "Bengaluru, India", "Jan 2025 – Sep 2026")
    add_bullet("Delivered domestic-scheme processing across the UAE (Jaywan), KSA, RuPay, CUP, PayPak, and MyDebit on CyberSource using Payments 2.0 microservices and Kafka event-driven settlement, contributing $30M in net revenue.")
    add_bullet("Embedded Claude-based agentic AI toolchains and AI-assisted workflows to accelerate solution architecture, drafting PRDs, acceptance criteria, and defect analyses—cutting manual delivery effort by 50%.")
    add_bullet("Accelerated cross-border integration delivery for SMCC Japan (EDQP, PIF, force-refund) using an agentic toolchain to automate specification validation and edge-case resolution.")
    add_bullet("Automated CAS self-service testing and settlement-file retrieval for acquirers, merchants, and technology partners, reducing file turnaround from days to under 24 hours and saving 8 FTEs of effort per sprint.")
    add_bullet("Shifted settlement generation from a centralized 400+ file per release model to on-demand, scoped partner access, eliminating release churn and coordination bottlenecks.")
    add_bullet("Defined Kafka-based REST event contracts with framework and orchestration teams for real-time transaction processing and settlement feeds.")

    # Visa TPM
    add_job_header("VISA", "Technical Product Manager – Visa Platform Connect (VPC)", "Bengaluru, India", "Mar 2024 – Dec 2024")
    add_bullet("Simplified technical implementation requirements for the Visa Instalment Plan in CEMEA, driving an 18% increase in regional transaction business volume.")
    add_bullet("Resolved critical production settlement incidents across CEMEA, APAC, Europe, and the Americas, introducing automated regeneration workflows that contributed to $9.2M in revenue protection.")
    add_bullet("Streamlined Credential-on-File (COF), Merchant-Initiated (MIT), and Customer-Initiated (CIT) tokenization flows for European partners, lifting token-based volume by 25%.")
    add_bullet("Led Business Enhancement Request (BER) release coordination across 100+ partner institutions, managing customer communications, cross-functional war rooms, and rollback planning.")

    # Visa TPA
    add_job_header("VISA", "Technical Product Analyst – Visa Platform Connect (VPC)", "Bengaluru, India", "Oct 2021 – Feb 2024")
    add_bullet("Supported core payment processing for Visa Platform Connect, an acquiring platform serving 100+ banking institutions across clearing, settlement, and transaction routing.")
    add_bullet("Collaborated with engineering and UX teams to scale partner onboarding to 400+ financial institutions through batch attribute updates and automated workflow enhancements.")
    add_bullet("Designed domestic co-brand transaction routing and BIN lookup logic for EFTPOS and MADA schemes.")

    # Cognizant Team Lead
    add_job_header("COGNIZANT", "Team Lead – Product Managers (Fiserv / First Data Account)", "Kolkata, India", "Jul 2019 – Sep 2021")
    add_bullet("Led a 12-member team of product managers and analysts supporting enterprise financial platforms, with full accountability for KPI/SLA tracking, stakeholder updates, and release governance.")
    add_bullet("Introduced upstream transaction validation in payment flows for enterprise merchants, helping protect $11M in revenue prior to major production release cycles.")

    # Cognizant PM
    add_job_header("COGNIZANT", "Product Manager – Transaction Validation Automation (First Data Account)", "Chennai, India", "Jul 2017 – Jun 2019")
    add_bullet("Launched an automated transaction-validation engine translating 13 technical specifications into algorithmic logic using RPA, Java, ASP.NET, and SQL, improving certification project cycles by 57%.")
    add_bullet("Managed sprint backlog, user stories, acceptance criteria, and A/B testing in an agile delivery model, saving 10 FTEs of manual effort and contributing to a 32% year-end revenue increase.")

    # Cognizant Cert PM
    add_job_header("COGNIZANT", "Certification Project Manager / Payments SME (First Data Account)", "Kolkata, India", "Feb 2014 – Jun 2017")
    add_bullet("Managed end-to-end certification and onboarding across 100+ enterprise clients; recognized as Cards & Payments Subject Matter Expert across ISO, EDC, and network clearing formats.")

    # --- TECHNICAL SKILLS ---
    add_section_header("Technical Skills & Tools")
    tech_skills = [
        ("Agile & Product Delivery: ", "Jira, Confluence, Backlog Prioritization, User Story & Acceptance Criteria Definition, PRD Creation, CSPO, Scrum, Kanban"),
        ("Banking, Payments & Schemes: ", "UAE Jaywan, Visa, Mastercard, ISO 8583, BASE I & BASE II, TC33.A, Tokenization, Acquirer Processing, Clearing & Settlement"),
        ("AI & Modern Platforms: ", "Agentic AI Toolchains, Claude, Gemini, OpenAI, Microservices, REST APIs, Kafka Event-Driven Systems, Automation Anywhere")
    ]
    for cat, items in tech_skills:
        add_bullet(items, bold_prefix=cat)

    # --- EDUCATION & CERTIFICATIONS ---
    add_section_header("Education & Certifications")
    p_ec = doc.add_paragraph()
    p_ec.paragraph_format.space_before = Pt(2)
    p_ec.paragraph_format.space_after = Pt(2)
    p_ec.paragraph_format.line_spacing = 1.12
    r_e1 = p_ec.add_run("B.Tech – Information Technology | West Bengal University of Technology (2012) | GPA: 8.6\n")
    r_e1.font.size = Pt(9)
    r_e2 = p_ec.add_run("Certifications: Certified Scrum Product Owner (CSPO, Scrum Alliance) | Product Lifecycle Management (PLM, Visa) | VPL Core Product (Visa)")
    r_e2.font.size = Pt(9)
    r_e2.font.color.rgb = RGBColor(71, 85, 105)

    doc.save(output_path)
    print(f"CBD Tailored CV saved to: {output_path}")

if __name__ == "__main__":
    out1 = r"c:\Users\BHASWATI\Downloads\binary\job-agent-starter\job-agent\applications\cbd-product-manager\Avijit_Mukherjee_CBD_Product_Manager.docx"
    build_cbd_tailored_cv(out1)

