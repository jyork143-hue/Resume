import os
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def add_heading(doc, text):
    heading = doc.add_heading(text, level=1)
    run = heading.runs[0]
    run.font.name = 'Arial'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(30, 58, 138)  # Executive Navy (#1e3a8a)
    heading.paragraph_format.space_before = Pt(14)
    heading.paragraph_format.space_after = Pt(4)

def add_job(doc, title, company, location, date, descriptions):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run_comp = p.add_run(f"{company} ")
    run_comp.bold = True
    run_comp.font.size = Pt(10.5)
    
    if location:
        run_loc = p.add_run(f"| {location}")
        run_loc.font.size = Pt(10.5)
    
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_after = Pt(4)
    run_title = p2.add_run(title)
    run_title.italic = True
    run_title.font.size = Pt(10.5)
    
    run_date = p2.add_run(f" ({date})")
    run_date.font.size = Pt(10.5)

    for desc in descriptions:
        p_desc = doc.add_paragraph(style='List Bullet')
        p_desc.paragraph_format.left_indent = Inches(0.25)
        p_desc.paragraph_format.space_after = Pt(3)
        run_d = p_desc.add_run(desc)
        run_d.font.size = Pt(10)

def main():
    doc = docx.Document()

    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)

    # Header / Contact Info
    header_p = doc.add_paragraph()
    header_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    header_p.paragraph_format.space_after = Pt(4)
    name_run = header_p.add_run("JOHNATHAN D. YORK, PMP, PMI-CPMAI\n")
    name_run.bold = True
    name_run.font.size = Pt(18)
    name_run.font.color.rgb = RGBColor(15, 23, 42)

    sub_run = header_p.add_run("Senior Program & Technical Program Leader | AI Transformation | Regulated Operations\n")
    sub_run.bold = True
    sub_run.font.size = Pt(10.5)
    sub_run.font.color.rgb = RGBColor(30, 58, 138)

    contact_run = header_p.add_run(
        "Maryland / Washington DC Metro | (843) 754-7350 | Jyork143@gmail.com | linkedin.com/in/johnathan-york\n"
        "C2C / Advisory (JD AI Solutions LLC): JDSolutionsAI.com | JDAISolutions.AI | contact@JDSolutionsAI.com"
    )
    contact_run.font.size = Pt(9.5)
    contact_run.font.color.rgb = RGBColor(71, 85, 105)

    # Executive Positioning & Profile
    add_heading(doc, "EXECUTIVE PROFILE")
    p_pos = doc.add_paragraph(
        "I lead complex programs where technology, operations, governance, and risk intersect. "
        "Open to senior contract engagements either through JD AI Solutions LLC on a corp-to-corp basis or through W-2 contract placement where required. "
        "Primary focus: Senior Program Manager, Technical Program Manager, AI Program Manager, PMO leadership, and regulated-industry transformation."
    )
    p_pos.paragraph_format.space_after = Pt(6)

    p_profile = doc.add_paragraph(
        "Senior Program and Technical Program Leader with over 15 years of progressive experience driving large-scale capital projects, "
        "engineering operations, and infrastructure modernization across heavily regulated nuclear, energy, and enterprise logistics architectures. "
        "U.S. Navy Nuclear Program veteran with deep mastery of nuclear plant operations, safety culture, and quality governance frameworks (JFMM, PPEP). "
        "Certified PMP and PMI-CPMAI with a proven track record managing $50M+ operational budgets, cross-functional technical teams, and governed AI "
        "and workflow automation integrations in mission-critical environments."
    )

    # Core Competencies
    add_heading(doc, "CORE CAPABILITIES")
    competencies = [
        "Program & Technical Program Management (PMP)",
        "AI Governance & Responsible AI Delivery (PMI-CPMAI)",
        "Capital Project Planning & Execution Process (PPEP)",
        "Regulated Nuclear Governance & Quality Assurance (JFMM)",
        "Multi-Million Dollar Budget & Earned Value Management",
        "GenAI Workflow Integration & Microsoft Copilot Deployment",
        "Root Cause Analysis, Risk Mitigation & Vendor Oversight",
        "High-Voltage Infrastructure, BESS & Mission-Critical Operations"
    ]
    
    table = doc.add_table(rows=4, cols=2)
    table.autofit = True
    for i in range(4):
        p1 = table.cell(i, 0).paragraphs[0]
        p1.style = 'List Bullet'
        p1.paragraph_format.left_indent = Inches(0.2)
        p1.add_run(competencies[i*2]).font.size = Pt(9.5)
        
        p2 = table.cell(i, 1).paragraphs[0]
        p2.style = 'List Bullet'
        p2.paragraph_format.left_indent = Inches(0.2)
        p2.add_run(competencies[i*2+1]).font.size = Pt(9.5)
    
    # Professional Experience
    add_heading(doc, "PROFESSIONAL EXPERIENCE")

    add_job(doc, "Founder & Principal Consultant", "JD AI SOLUTIONS LLC", "Maryland, USA", "2026 - Present", [
        "Founded and launched a Maryland-licensed AI governance, secure automation, and sensitive-data boundary solutions company (SBA-Certified SDVOSB / VOSB) helping regulated and operationally complex organizations move from AI interest to accountable implementation.",
        "Leveraging 15+ years of program and operational leadership experience to deliver strategic project execution, Corp-to-Corp technical program leadership, and digital transformation frameworks across regulated and mission-critical industries."
    ])

    add_job(doc, "Nuclear Project Manager, Capital Projects", "XCEL ENERGY (via SelectSource Int'l)", "Prairie Island Nuclear Plant, MN", "2025 - 2025", [
        "Led capital project planning and project execution support under the strict, regulated parameters of a commercial nuclear power facility.",
        "Managed complex project workflows, risk matrices, and data governance alignments supporting the site's Project Planning and Execution Process (PPEP).",
        "Evaluated heavy engineering and operational workflows to isolate cognitive vs. non-cognitive tasks, mapping optimal entry points for data-driven tools to enhance team productivity and decision support while maintaining strict safety overhead.",
        "Pioneered AI adoption frameworks and governed training deployment for Microsoft Copilot, enabling legacy technical teams to safely leverage advanced computing within regulatory boundaries and limitations.",
        "Evaluated emerging technologies against rigorous corporate risk, data privacy, and operational readiness metrics utilizing foundational PMI-CPMAI™ frameworks."
    ])

    add_job(doc, "Engineering Project Manager, Intelligent Tools", "STARBUCKS TRYER LABORATORY (via Ascendion)", "Seattle, WA", "2024 - 2025", [
        "Spearheaded the full-lifecycle development, risk management, and strategic deployment of data-centric tools and custom automation applications to optimize enterprise asset and infrastructure tracking.",
        "Deployed a centralized Maintenance Management System and automated Inventory Tracking Tool that increased real-time operational visibility by 80% and reduced YoY corporate audit gaps by 45%.",
        "Developed customized electrical distribution and infrastructure telemetry tools that accelerated time-to-deployment by 30% across cross-functional facilities.",
        "Enforced rigid technology governance and validation frameworks, establishing explicit corporate metrics for risk detection, data privacy compliance, transparency, and data integrity."
    ])

    add_job(doc, "Program Manager, North America Operational Excellence", "AMAZON", "North America", "2021 - 2024", [
        "Directed overarching project installation, critical infrastructure logistics, and network operations for complex automated logistics networks across North America.",
        "Managed international resource allocations, vendor contracts, and risk response strategies for a $50M annual operations budget, delivering programmatic cost savings that outperformed target goals by 42%.",
        "Automated comprehensive asset analysis and predictive health metrics for critical regional electrical power and network connectivity systems, reducing product downtime by 40%.",
        "Utilized cloud-powered digital augmentations and modernized fleet mapping systems to drive a 37% increase in hardware infrastructure activations while cutting installation failure rates by 52.1% YoY.",
        "Revamped workflow automation pipelines to eliminate operational backlogs, lifting on-site delivery and scheduling accuracy by 23% YoY to align with enterprise performance metrics."
    ])

    add_job(doc, "Site Operations Manager, Flagship Test Center", "PV EVOLUTION LABS / KIWA PVEL", "South San Francisco, CA", "2019 - 2021", [
        "Recruited to greenfield-onboard, staff, and execute operations for a high-voltage testing facility, managing its strategic scale-up into a multi-megawatt Inverter and Battery Energy Storage Systems (BESS) lab to anchor flagship Product Qualification Programs (PQP).",
        "Commissioned and supervised the complex engineering installation of two separate independent AC grid simulators, facilitating empirical bankability testing under simulated grid anomalies and dynamic utility conditions.",
        "Directed daily operations for extended-reliability testing, including environmental chamber testing, Incidence Angle Modifier (IAM), and High-Throughput Module Inspections (HTMI) under rigid safety standards.",
        "Maintained 100% operational continuity and testing throughput while managing laboratory personnel through the height of the COVID-19 pandemic and a major international corporate acquisition by Kiwa."
    ])

    add_job(doc, "Project Lead & Autonomous Systems Researcher", "NASA STUDENT LAUNCH PROGRAM / UNIVERSITY OF HAWAII", "Honolulu, HI", "2017 - 2019", [
        "Concepted, engineered, and built a fully autonomous, solar-powered smart irrigation system designed to restore native vegetation and mitigate critical soil erosion on the island of Kaho'olawe; presented findings at the national SACNAS Conference (San Antonio, TX).",
        "Formulated and executed real-time Python-based control algorithms and telemetry logic for autonomous platforms, successfully decreasing system latency by 40% and ensuring 100% prototype field reliability.",
        "Led competitive engineering cohorts in the international ARLISS aerospace competition in Black Rock, Nevada, commanding the design and deployment of autonomous satellites under 100% NASA flight-readiness compliance."
    ])

    add_job(doc, "Nuclear Electrician / Submarine Warfare Specialist", "UNITED STATES NAVY | Submarine Force", "USS Columbia (SSN 771) & MARF Prototype", "2009 - 2016", [
        "Supervised, maintained, and operated mission-critical nuclear propulsion and electrical generation plant infrastructure under strict Department of Energy (DOE) and military operational tolerances.",
        "Selected as a Junior Staff Instructor at the MARF DOE nuclear training prototype based on technical aptitude; qualified as a Shutdown Reactor Operator (SRO), Throttleman, Craftsman, and Dual-Plant Cleanliness Inspector (Reactor/Secondary), training future fleet operators.",
        "Partnered directly with the Engineering Department Master Chief to execute a comprehensive operational revamp of the submarine's engineering department; qualified as Quality Assurance Inspector (QAI), QA Planner (QAP), and QA Work Center Supervisor under the rigid structural governance of the Joint Fleet Maintenance Manual (JFMM).",
        "Served as the primary Ship's Force Representative for complex shipyard overhaul periods due to expertise in auditing and executing highly regulated technical work packages and technical requirements.",
        "Crew recipient of the Arleigh Burke Fleet Trophy, Battle Effectiveness Award (Battle \"E\"), and Engineering Excellence Award (Engineering \"E\") across multiple classified national security deployments and a maximum-grade Operational Reactor Safeguards Examination (ORSE).",
        "Earned the U.S. Navy Submarine Warfare Designation (\"Dolphins\") in an unprecedented 3 months, dramatically outperforming the standard 12-to-18-month qualification timeline."
    ])

    # Education & Credentials
    add_heading(doc, "EDUCATION & CREDENTIALS")
    
    edu = [
        "University of Maryland Global Campus - Bachelor of Science in AI Applications (In Progress)",
        "University of Maryland Global Campus - Associate of Arts in Business and Management",
        "US Naval Nuclear Power School - Graduate (Nuclear Power Plant Construction & Reactor Theory)",
        "US Naval Nuclear Training Prototype - Graduate (Nuclear Plant Operation & Maintenance)"
    ]
    for e in edu:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(e).font.size = Pt(10)

    p_cert_title = doc.add_paragraph("\nCertifications & Professional Credentials:")
    p_cert_title.runs[0].bold = True
    
    certs = [
        "Project Management Professional (PMP)® - Project Management Institute (Credential ID: 3959827)",
        "PMI Certified Professional in Managing AI (PMI-CPMAI)™ - Project Management Institute",
        "PMI Advanced Micro-Credentials:"
    ]
    for c in certs:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(c).font.size = Pt(10)

    sub_certs = [
        "Practical Application of GenAI for Project Managers",
        "AI in Infrastructure & Construction Management",
        "Talking to AI (Prompt Engineering)",
        "Introduction to Cognitive Project Management in AI"
    ]
    for sc in sub_certs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.5)
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"- {sc}").font.size = Pt(10)

    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Johnathan_York_Resume.docx")
    doc.save(output_path)

if __name__ == "__main__":
    main()
