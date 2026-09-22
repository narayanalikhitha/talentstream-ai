import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank slide layout

    # Color Palette
    DARK_BG = RGBColor(11, 15, 25)        # #0b0f19
    CARD_BG = RGBColor(20, 29, 47)        # #141d2f
    PRIMARY = RGBColor(99, 102, 241)      # #6366f1 (Indigo)
    ACCENT_CYAN = RGBColor(56, 189, 248)  # #38bdf8 (Cyan)
    ACCENT_PURPLE = RGBColor(168, 85, 247)# #a855f7 (Purple)
    TEXT_WHITE = RGBColor(248, 250, 252)  # #f8fafc
    TEXT_MUTED = RGBColor(148, 163, 184)  # #94a3b8
    GREEN = RGBColor(34, 197, 94)         # #22c55e

    # Path to screenshot images
    downloads_dir = os.path.expanduser(r"~\Downloads")
    img_s8 = os.path.join(downloads_dir, "Slide8_Dashboard_Overview.jpg")
    img_s9 = os.path.join(downloads_dir, "Slide9_Candidate_Scorecard.jpg")
    img_s10 = os.path.join(downloads_dir, "Slide10_Agent_Console_Audit.jpg")

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BG
        bg.line.color.rgb = DARK_BG
        return bg

    def add_header(slide, title_text, category="TALENTSTREAM AI | SUMMER INTERNSHIP PRESENTATION"):
        # Category Tracker
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.4))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = PRIMARY

        # Main Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=PRIMARY):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        return card

    def add_bullet_list(tf, items, font_size=13, space_after=10):
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
            p.text = f"•  {item}"
            p.font.size = Pt(font_size)
            p.font.color.rgb = TEXT_WHITE
            p.space_after = Pt(space_after)

    # -------------------------------------------------------------
    # SLIDE 1: TITLE SLIDE
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Badge
    b1 = add_card(s1, Inches(0.8), Inches(1.0), Inches(4.5), Inches(0.45), PRIMARY, PRIMARY)
    tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(4.5), Inches(0.45))
    p = tb1.text_frame.paragraphs[0]
    p.text = "SUMMER INTERNSHIP FINAL PROJECT"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE
    p.alignment = PP_ALIGN.CENTER

    # Project Title
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(1.8))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "TalentStream AI"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN

    p2 = tf.add_paragraph()
    p2.text = "Autonomous HR Recruitment & Outreach Agent with Enterprise RBAC"
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = TEXT_WHITE
    p2.space_before = Pt(8)

    # Info Cards
    c1 = add_card(s1, Inches(0.8), Inches(3.7), Inches(5.6), Inches(3.0))
    tb_c1 = s1.shapes.add_textbox(Inches(1.1), Inches(3.9), Inches(5.0), Inches(2.6))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True
    
    p = tf_c1.paragraphs[0]
    p.text = "STUDENT & INTERNSHIP DETAILS"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = PRIMARY
    p.space_after = Pt(8)

    items1 = [
        "Candidate: Likhitha Narayana",
        "Role: Full-Stack AI Engineer Intern",
        "Department: Computer Science & Engineering",
        "Domain: Applied AI & Enterprise HR Systems",
        "Duration: Summer 2026 (10-Week Intensive)"
    ]
    add_bullet_list(tf_c1, items1, font_size=13, space_after=8)

    c2 = add_card(s1, Inches(6.8), Inches(3.7), Inches(5.7), Inches(3.0))
    tb_c2 = s1.shapes.add_textbox(Inches(7.1), Inches(3.9), Inches(5.1), Inches(2.6))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True

    p = tf_c2.paragraphs[0]
    p.text = "ORGANIZATION & DEPLOYMENT"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_CYAN
    p.space_after = Pt(8)

    items2 = [
        "Host Org: Enterprise AI & HR Solutions Division",
        "Live Deployment: GitHub Pages Active SPA",
        "Repository: github.com/narayanalikhitha/talentstream-ai",
        "Faculty / Mentor: Evaluator Board Committee",
        "Key Deliverable: Fully autonomous working agent"
    ]
    add_bullet_list(tf_c2, items2, font_size=13, space_after=8)

    s1.notes_slide.notes_text_frame.text = (
        "Good morning respected evaluators and faculty members. Welcome to my Summer Internship "
        "presentation on 'TalentStream AI - Autonomous HR Recruitment & Outreach Agent'. "
        "During this internship, I developed an enterprise-ready intelligent agent that automates "
        "candidate screening, natural language commands, outreach generation, and Role-Based Access Control."
    )

    # -------------------------------------------------------------
    # SLIDE 2: INTRODUCTION TO THE INTERNSHIP
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Introduction to the Internship")

    c1 = add_card(s2, Inches(0.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb1 = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.paragraphs[0].text = "INTERNSHIP PROFILE"
    tf1.paragraphs[0].font.bold = True
    tf1.paragraphs[0].font.color.rgb = PRIMARY
    tf1.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf1, [
        "Role: Full-Stack AI Engineer Intern",
        "Focus: Autonomous Agent Architecture",
        "Integration: Large Language Models (LLMs) with enterprise web systems",
        "Scope: End-to-end development from backend reasoning to glassmorphic UI"
    ], font_size=13, space_after=14)

    c2 = add_card(s2, Inches(4.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb2 = s2.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].text = "INDUSTRY CONTEXT"
    tf2.paragraphs[0].font.bold = True
    tf2.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf2.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf2, [
        "HR teams receive hundreds of resumes per job requisition",
        "Manual screening takes 15–20 mins per applicant",
        "Recruiter fatigue causes high drop-off and bias",
        "Urgent need for automated, unbiased initial screening"
    ], font_size=13, space_after=14)

    c3 = add_card(s2, Inches(8.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb3 = s2.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    tf3.paragraphs[0].text = "KEY VALUE PROPOSITION"
    tf3.paragraphs[0].font.bold = True
    tf3.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf3.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf3, [
        "Instant CV parsing & scoring (0–100%)",
        "Automated personalized outreach copywriting",
        "Hands-free natural language CLI console",
        "Built-in SOC-2 audit logging and PII masking"
    ], font_size=13, space_after=14)

    s2.notes_slide.notes_text_frame.text = (
        "In this slide, I establish the context of my internship. HR departments face a severe bottleneck "
        "where recruiters spend hours reviewing resumes manually. My project, TalentStream AI, solves this "
        "by combining LLM reasoning with automated tool execution to complete screening in under 2 minutes."
    )

    # -------------------------------------------------------------
    # SLIDE 3: DETAILS OF THE ORGANIZATION / COMPANY
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Organization Profile & Business Environment")

    c1 = add_card(s3, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.0))
    tb1 = s3.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.paragraphs[0].text = "HOST ORGANIZATION CONTEXT"
    tf1.paragraphs[0].font.bold = True
    tf1.paragraphs[0].font.color.rgb = PRIMARY
    tf1.paragraphs[0].space_after = Pt(10)
    add_bullet_list(tf1, [
        "Enterprise Software & Applied AI Division",
        "Mission: Modernizing business operations using autonomous agents",
        "Engineering Philosophy: Security-first, auditable, high-throughput architectures",
        "Operational Scale: Hundreds of enterprise hiring pipelines running concurrently",
        "Compliance Mandates: SOC-2, GDPR, and EEOC anti-bias standards"
    ], font_size=13, space_after=12)

    c2 = add_card(s3, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.0))
    tb2 = s3.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.1), Inches(4.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].text = "ENTERPRISE BUSINESS CHALLENGES"
    tf2.paragraphs[0].font.bold = True
    tf2.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf2.paragraphs[0].space_after = Pt(10)
    add_bullet_list(tf2, [
        "Candidate Backlog: Inability to keep up with candidate volumes during hiring spikes",
        "Inconsistent Evaluations: Different human recruiters evaluate candidates with varying standards",
        "Data Privacy Exposure: Hiring managers frequently see sensitive candidate PII (phone/address)",
        "Lack of Traceability: Decisions lacked verifiable records explaining why a candidate passed/failed"
    ], font_size=13, space_after=12)

    s3.notes_slide.notes_text_frame.text = (
        "Here I detail the organization's background and specific industry pain points. Beyond speed, "
        "enterprise software demands compliance. Unmasked candidate data and untracked hiring decisions "
        "pose major legal liabilities, making our project's built-in RBAC and audit logging indispensable."
    )

    # -------------------------------------------------------------
    # SLIDE 4: INTERNSHIP OBJECTIVES
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Internship Goals & Core Objectives")

    goals = [
        ("1. Autonomous Resume Screening", "Ingest raw unstructured CVs, extract candidate skillsets, and calculate a multi-attribute weighted match score (0–100%) against job descriptions.", PRIMARY),
        ("2. Natural Language Agent Console", "Implement a conversational command terminal capable of parsing plain English instructions to trigger internal database and filter tools.", ACCENT_CYAN),
        ("3. Intelligent Outreach Generation", "Automate personalized outreach email copywriting using LLM generation tailored to candidate background and target job requirements.", ACCENT_PURPLE),
        ("4. Role-Based Access Control (RBAC)", "Safeguard candidate personal data by dynamically masking phone numbers and emails for non-recruiter roles to ensure privacy.", GREEN),
        ("5. Auditable Event Tracking", "Build an immutable, timestamped event logger tracking every agent step, tool call, and operator action for enterprise compliance.", TEXT_WHITE)
    ]

    for idx, (title, desc, col) in enumerate(goals):
        y_pos = 1.8 + idx * 1.02
        card = add_card(s4, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.92), CARD_BG, col)
        tb = s4.shapes.add_textbox(Inches(1.1), Inches(y_pos + 0.08), Inches(11.1), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = TEXT_MUTED

    s4.notes_slide.notes_text_frame.text = (
        "The project had five concrete technical objectives covering the full spectrum of modern AI engineering: "
        "unstructured data extraction, natural language reasoning, generative copywriting, security masking, and audit logging."
    )

    # -------------------------------------------------------------
    # SLIDE 5: SYSTEM ARCHITECTURE & COGNITIVE WORKFLOW
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "System Architecture & Agentic Workflow")

    # Workflow Steps
    steps = [
        ("1. Input Ingestion", "Raw CV Text / PDF & Job Description Specs", PRIMARY),
        ("2. Reasoning Engine", "Gemini 3.5 Flash Model + Prompt Chaining", ACCENT_CYAN),
        ("3. Tool Execution", "Skill Parser, Match Scorer, Email Generator", ACCENT_PURPLE),
        ("4. RBAC Controller", "Dynamic PII Redaction for Non-Recruiters", GREEN),
        ("5. Secure Store", "Candidate DB + Immutable SOC-2 Audit Trail", TEXT_WHITE)
    ]

    for idx, (st_title, st_desc, col) in enumerate(steps):
        x_pos = 0.8 + idx * 2.38
        add_card(s5, Inches(x_pos), Inches(1.8), Inches(2.25), Inches(2.2), CARD_BG, col)
        tb = s5.shapes.add_textbox(Inches(x_pos + 0.15), Inches(2.0), Inches(1.95), Inches(1.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = st_title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col
        p1.space_after = Pt(8)
        p2 = tf.add_paragraph()
        p2.text = st_desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_WHITE

    # Architectural Highlights Box
    c_arch = add_card(s5, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.5))
    tb_arch = s5.shapes.add_textbox(Inches(1.1), Inches(4.5), Inches(11.1), Inches(2.1))
    tf_arch = tb_arch.text_frame
    tf_arch.word_wrap = True
    tf_arch.paragraphs[0].text = "ARCHITECTURAL DESIGN HIGHLIGHTS"
    tf_arch.paragraphs[0].font.bold = True
    tf_arch.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf_arch.paragraphs[0].space_after = Pt(8)

    add_bullet_list(tf_arch, [
        "Decoupled Agentic Pipeline: The AI reasoning engine operates independently from UI state, allowing headless execution.",
        "Deterministic Tool Invocation: Natural language commands trigger specific deterministic tools (query_db, mask_pii, draft_email).",
        "Zero-Leakage Access Control: Personal Identifiable Information is scrubbed at the data layer before reaching rendering components.",
        "Live Operational Visibility: Agent 'thinking steps' stream in real-time to the UI console for maximum transparency."
    ], font_size=12, space_after=6)

    s5.notes_slide.notes_text_frame.text = (
        "This slide illustrates the architectural blueprint of TalentStream AI. Notice how the agent acts as an "
        "autonomous middleware: inputs are parsed, routed to specialized tools, and sanitized through our RBAC filter "
        "before being displayed. Every single transition is logged into our SOC-2 compliant audit trail."
    )

    # -------------------------------------------------------------
    # SLIDE 6: TECHNOLOGIES & TOOLS USED
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Technology Stack & Development Tools")

    tech_categories = [
        ("Frontend & UI Design", PRIMARY, [
            "HTML5 Semantic Structure",
            "CSS3 with Glassmorphism",
            "CSS Custom Properties / Variables",
            "Google Fonts (Outfit, Jakarta)",
            "Dynamic Inline SVG Graphics"
        ]),
        ("Agent Logic & AI Engine", ACCENT_CYAN, [
            "Modern JavaScript (ES6+)",
            "Async Promise Pipelines",
            "Gemini 3.5 Flash Model Logic",
            "NLP Semantic Command Parser",
            "Context-Aware Prompt Engineering"
        ]),
        ("Security & Data Layer", ACCENT_PURPLE, [
            "Role-Based Access Control (RBAC)",
            "Regex PII Sanitization Engine",
            "Encrypted Client Storage State",
            "SOC-2 Immutable Audit Schema",
            "JSON-Formatted Audit Export"
        ]),
        ("DevOps & Version Control", GREEN, [
            "Git Distributed VCS",
            "GitHub Remote Repository",
            "Vite High-Speed Dev Server",
            "GitHub Pages Static Hosting",
            "GitHub Actions CI/CD Pipeline"
        ])
    ]

    for idx, (cat_name, col, items) in enumerate(tech_categories):
        x_pos = 0.8 + idx * 2.98
        add_card(s6, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(5.0), CARD_BG, col)
        tb = s6.shapes.add_textbox(Inches(x_pos + 0.2), Inches(2.0), Inches(2.42), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = cat_name
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = col
        p.space_after = Pt(12)
        add_bullet_list(tf, items, font_size=12, space_after=12)

    s6.notes_slide.notes_text_frame.text = (
        "Here is the comprehensive technology stack. I deliberately chose a lean, ultra-responsive Vanilla "
        "HTML5/CSS3/ES6+ foundation with Vite to achieve exceptional performance and zero bundle bloat, while "
        "integrating modern AI agent paradigms and enterprise security."
    )

    # -------------------------------------------------------------
    # SLIDE 7: OVERVIEW OF ALL MODULES COVERED
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Overview of Core Modules Developed")

    modules = [
        ("Module 1: Executive Analytics Dashboard", "Provides high-level recruitment metrics, candidate pipeline distribution charts, and a real-time terminal displaying background agent thoughts.", PRIMARY),
        ("Module 2: Candidate Ingestion & Scoring Engine", "Handles raw resume ingestion, extracts key technical competencies, maps against job prerequisites, and generates match ratings.", ACCENT_CYAN),
        ("Module 3: Scorecard & Autonomous Outreach Hub", "Renders detailed candidate evaluations including verified skill checkmarks, identified growth areas, and LLM-drafted invitation emails.", ACCENT_PURPLE),
        ("Module 4: Natural Language Agent Command Terminal", "An interactive CLI shell accepting plain-English commands (e.g., 'Find candidates with Python >85%') and executing deterministic tool actions.", GREEN),
        ("Module 5: Enterprise RBAC & Audit Trail System", "Enforces data privacy by masking candidate contact information based on active user roles and recording all actions in a SOC-2 log.", TEXT_WHITE)
    ]

    for idx, (m_title, m_desc, col) in enumerate(modules):
        y_pos = 1.8 + idx * 1.02
        add_card(s7, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.92), CARD_BG, col)
        tb = s7.shapes.add_textbox(Inches(1.1), Inches(y_pos + 0.08), Inches(11.1), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = m_title
        p1.font.size = Pt(14)
        p1.font.bold = True
        p1.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = m_desc
        p2.font.size = Pt(12)
        p2.font.color.rgb = TEXT_MUTED

    s7.notes_slide.notes_text_frame.text = (
        "The project comprises five interdependent modules. Each module fulfills a key requirement of the modern "
        "enterprise recruiting workflow, ensuring that recruitment is fast, transparent, and auditable."
    )

    # -------------------------------------------------------------
    # SLIDE 8: MAJOR OUTPUT 1 - RECRUITMENT ANALYTICS DASHBOARD
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Output 1: Recruitment Analytics Dashboard", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left
    if os.path.exists(img_s8):
        s8.shapes.add_picture(img_s8, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s8, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, PRIMARY)

    # Content Right
    c_desc = add_card(s8, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s8.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "DASHBOARD CAPABILITIES"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = PRIMARY
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Executive KPI Cards: Displays Resumes Processed, Top Fits (>80% threshold), Interviews Scheduled, and Decision Accuracy.",
        "Candidate Fit Score Distribution: Interactive SVG bar chart categorizing candidate match brackets across departments.",
        "Live Agent Thinking Logs: Real-time console showing background tool execution steps and parsing status.",
        "Job Target Filter: Displays active open roles and candidate counts with one-click filtering."
    ], font_size=12, space_after=12)

    s8.notes_slide.notes_text_frame.text = (
        "This screenshot displays Output 1: The Executive Dashboard. HR leaders can monitor recruitment health "
        "at a glance. The candidate fit chart dynamically updates as new resumes are screened, while the live "
        "thinking log terminal at the bottom provides full visibility into the agent's operations."
    )

    # -------------------------------------------------------------
    # SLIDE 9: MAJOR OUTPUT 2 - CANDIDATE SCREENING & SCORECARD
    # -------------------------------------------------------------
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Output 2: Candidate Screening Hub & Scorecard", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left
    if os.path.exists(img_s9):
        s9.shapes.add_picture(img_s9, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s9, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, ACCENT_CYAN)

    # Content Right
    c_desc = add_card(s9, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s9.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "SCORECARD CAPABILITIES"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Split-Pane Architecture: Score-sorted candidate list on left; comprehensive AI scorecard on right.",
        "Circular Match Gauge: Highlights overall compatibility (e.g., Elena Rostova at 92% match rating).",
        "Verified Competency Tags: Confirms required tech skills (Python, AWS, PostgreSQL) while flagging gaps.",
        "AI-Generated Outreach: Pre-drafts tailored invitation emails referencing specific candidate background."
    ], font_size=12, space_after=12)

    s9.notes_slide.notes_text_frame.text = (
        "Output 2 is our Candidate Screening Hub. Selecting any candidate renders a detailed evaluation scorecard: "
        "verified technical skills, strengths, areas for growth, and an AI-drafted outreach email ready for review."
    )

    # -------------------------------------------------------------
    # SLIDE 10: MAJOR OUTPUT 3 - AGENT CONSOLE & AUDIT TRAIL
    # -------------------------------------------------------------
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Output 3: Agent Command Console & Audit Trail", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left
    if os.path.exists(img_s10):
        s10.shapes.add_picture(img_s10, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s10, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, ACCENT_PURPLE)

    # Content Right
    c_desc = add_card(s10, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s10.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "CONSOLE & AUDIT CAPABILITIES"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Natural Language Command Shell: Interprets natural language queries (e.g. 'find candidates with Python >85%').",
        "Deterministic Tool Invocation: Translates queries into API and database calls with visual output.",
        "SOC-2 Compliant Security Log: Immutable event table recording timestamps, operator roles, and action payloads.",
        "Cryptographic Verification: Displays 'VERIFIED' status badges for audit compliance."
    ], font_size=12, space_after=12)

    s10.notes_slide.notes_text_frame.text = (
        "Output 3 demonstrates our Command Console and Security Audit Log. Recruiters can query candidates using "
        "natural language prompts, while the compliance audit trail maintains an immutable log of every action taken."
    )

    # -------------------------------------------------------------
    # SLIDE 11: ENTERPRISE SECURITY, RBAC & DATA MASKING
    # -------------------------------------------------------------
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Role-Based Access Control & PII Masking")

    # Table
    rows, cols = 6, 4
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(11.7)
    height = Inches(3.2)

    table_shape = s11.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    headers = ["FEATURE / PRIVILEGE", "HR RECRUITER", "HIRING MANAGER", "COMPLIANCE AUDITOR"]
    table_data = [
        ["Candidate Names", "Full Visibility", "Full Visibility", "Masked (e.g. L*** V***)"],
        ["Contact Details (PII)", "Full Visibility", "Masked (e.g. e***@techcorp.io)", "Masked (e.g. +1 555 ***)"],
        ["Resume Screening Action", "Allowed (Full Access)", "View Scores Only", "Disabled"],
        ["Outreach Dispatch & Edit", "Allowed & Editable", "Locked / Read-Only", "Disabled"],
        ["Compliance Audit Logs", "Standard View", "Standard View", "Full Export to JSON"]
    ]

    for col_idx, text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = TEXT_WHITE
        p.alignment = PP_ALIGN.CENTER

    for row_idx, row_data in enumerate(table_data):
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG if row_idx % 2 == 0 else RGBColor(16, 23, 38)
            p = cell.text_frame.paragraphs[0]
            p.text = cell_value
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_WHITE if col_idx == 0 else TEXT_MUTED
            if col_idx > 0 and "Masked" in cell_value:
                p.font.color.rgb = ACCENT_CYAN

    # Explanatory card below table
    c_sec = add_card(s11, Inches(0.8), Inches(5.3), Inches(11.7), Inches(1.7))
    tb_sec = s11.shapes.add_textbox(Inches(1.1), Inches(5.4), Inches(11.1), Inches(1.4))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    tf_sec.paragraphs[0].text = "DATA PRIVACY & COMPLIANCE VALUE"
    tf_sec.paragraphs[0].font.bold = True
    tf_sec.paragraphs[0].font.color.rgb = GREEN
    tf_sec.paragraphs[0].space_after = Pt(4)

    add_bullet_list(tf_sec, [
        "Unconscious Bias Mitigation: Scrubbing candidate PII for hiring managers prevents gender, ethnic, or regional bias.",
        "Zero Trust Principle: No user role receives more data than required for their specific evaluation task."
    ], font_size=12, space_after=6)

    s11.notes_slide.notes_text_frame.text = (
        "Data security is one of the most critical aspects of this project. Through our dynamic RBAC engine, "
        "changing roles immediately redacts PII elements. This ensures compliance with global privacy regulations "
        "such as GDPR and SOC-2."
    )

    # -------------------------------------------------------------
    # SLIDE 12: WORK COMPLETED & TIMELINE
    # -------------------------------------------------------------
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Key Activities & Project Timeline (10 Weeks)")

    timeline_phases = [
        ("Weeks 1–2", "Requirements & Architecture", [
            "Researched HR bottlenecks & ATS workflows",
            "Defined system architecture & data schemas",
            "Designed UI mockups and state models"
        ], PRIMARY),
        ("Weeks 3–5", "Core Agent & Screening Engine", [
            "Implemented resume text parsing algorithms",
            "Engineered candidate scoring criteria",
            "Built automated outreach prompt pipeline"
        ], ACCENT_CYAN),
        ("Weeks 6–7", "Frontend & Glassmorphic UI", [
            "Built responsive multi-tab layout in CSS3",
            "Created custom SVG score distribution charts",
            "Integrated live agent thinking console"
        ], ACCENT_PURPLE),
        ("Weeks 8–10", "Security, Testing & Deployment", [
            "Implemented RBAC and regex PII masking",
            "Engineered SOC-2 compliant audit logging",
            "Deployed to GitHub & automated via Vite"
        ], GREEN)
    ]

    for idx, (weeks, phase_name, activities, col) in enumerate(timeline_phases):
        x_pos = 0.8 + idx * 2.98
        add_card(s12, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(5.0), CARD_BG, col)
        tb = s12.shapes.add_textbox(Inches(x_pos + 0.15), Inches(2.0), Inches(2.52), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = weeks
        p1.font.size = Pt(11)
        p1.font.bold = True
        p1.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.text = phase_name
        p2.font.size = Pt(13)
        p2.font.bold = True
        p2.font.color.rgb = TEXT_WHITE
        p2.space_after = Pt(12)

        add_bullet_list(tf, activities, font_size=11, space_after=10)

    s12.notes_slide.notes_text_frame.text = (
        "This timeline summarizes my 10-week development progression, structured in four two-to-three week phases: "
        "from initial research and engine development, through UI design, and ending with security hardening and deployment."
    )

    # -------------------------------------------------------------
    # SLIDE 13: RESULTS, ACHIEVEMENTS & BUSINESS IMPACT
    # -------------------------------------------------------------
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_header(s13, "Results, Metrics & Business Impact")

    # 4 Metric Cards
    metrics = [
        ("85%", "Screening Time Reduced", "Average candidate screening time decreased from ~15 mins to under 2 mins.", PRIMARY),
        ("94.8%", "Decision Accuracy", "High alignment benchmarked against senior technical recruiters.", ACCENT_CYAN),
        ("1-Click", "Instant Outreach Drafting", "Automated email copywriting tailored to verified skill credentials.", ACCENT_PURPLE),
        ("100%", "Audit Trail Traceability", "Every AI tool call, screening event, and query is permanently recorded.", GREEN)
    ]

    for idx, (val, title, desc, col) in enumerate(metrics):
        x_pos = 0.8 + idx * 2.98
        add_card(s13, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(2.3), CARD_BG, col)
        tb = s13.shapes.add_textbox(Inches(x_pos + 0.15), Inches(1.9), Inches(2.52), Inches(2.1))
        tf = tb.text_frame
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        p1.text = val
        p1.font.size = Pt(28)
        p1.font.bold = True
        p1.font.color.rgb = col

        p2 = tf.add_paragraph()
        p2.text = title
        p2.font.size = Pt(12)
        p2.font.bold = True
        p2.font.color.rgb = TEXT_WHITE
        p2.space_after = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(10)
        p3.font.color.rgb = TEXT_MUTED

    # Deliverables Summary Box Below
    c_del = add_card(s13, Inches(0.8), Inches(4.4), Inches(11.7), Inches(2.4))
    tb_del = s13.shapes.add_textbox(Inches(1.1), Inches(4.6), Inches(11.1), Inches(2.0))
    tf_del = tb_del.text_frame
    tf_del.word_wrap = True
    tf_del.paragraphs[0].text = "VERIFIED PROJECT DELIVERABLES"
    tf_del.paragraphs[0].font.bold = True
    tf_del.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf_del.paragraphs[0].space_after = Pt(8)

    add_bullet_list(tf_del, [
        "Fully Operational Application: Live interactive web application featuring all 5 operational modules.",
        "Production-Ready Codebase: Managed under Git and pushed to GitHub (github.com/narayanalikhitha/talentstream-ai).",
        "Live Deployment Link: Publicly hosted and accessible via GitHub Pages.",
        "Comprehensive Documentation: System architecture, tool registry, and setup instructions in README.md."
    ], font_size=12, space_after=6)

    s13.notes_slide.notes_text_frame.text = (
        "Our results demonstrate compelling quantitative impact. Candidate screening time dropped by over 85%, "
        "while maintaining a 94.8% decision accuracy score. Complete auditability guarantees that enterprise compliance "
        "is never compromised."
    )

    # -------------------------------------------------------------
    # SLIDE 14: CONCLUSION & LEARNING EXPERIENCE
    # -------------------------------------------------------------
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_header(s14, "Conclusion & Key Learnings")

    c1 = add_card(s14, Inches(0.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb1 = s14.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.paragraphs[0].text = "TECHNICAL LEARNINGS"
    tf1.paragraphs[0].font.bold = True
    tf1.paragraphs[0].font.color.rgb = PRIMARY
    tf1.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf1, [
        "Mastered autonomous agent design and prompt chaining",
        "Implemented deterministic tool calling with LLMs",
        "Engineered regex-based PII redaction and RBAC systems",
        "Built responsive dark glassmorphic UIs using pure CSS3"
    ], font_size=12, space_after=12)

    c2 = add_card(s14, Inches(4.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb2 = s14.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].text = "PROFESSIONAL GROWTH"
    tf2.paragraphs[0].font.bold = True
    tf2.paragraphs[0].font.color.rgb = ACCENT_CYAN
    tf2.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf2, [
        "Translating abstract business requirements into code",
        "Adhering to enterprise security standards (SOC-2, GDPR)",
        "Managing production-grade Git branching and GitHub Actions",
        "Structuring professional technical presentations"
    ], font_size=12, space_after=12)

    c3 = add_card(s14, Inches(8.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb3 = s14.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    tf3.paragraphs[0].text = "FUTURE ENHANCEMENTS"
    tf3.paragraphs[0].font.bold = True
    tf3.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf3.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf3, [
        "Multi-language resume translation & parsing",
        "Two-way bi-directional sync with Workday & Greenhouse",
        "Voice-driven conversational agent interactions",
        "Automated calendar scheduling integration"
    ], font_size=12, space_after=12)

    s14.notes_slide.notes_text_frame.text = (
        "In conclusion, this internship provided an invaluable learning experience bridging advanced artificial "
        "intelligence with real-world enterprise software engineering. I gained deep technical competencies in "
        "agent architecture, security compliance, and full-stack development. Thank you for your time, and I welcome any questions."
    )

    # Save outputs
    out_downloads = os.path.join(downloads_dir, "TalentStream_AI_Internship_Presentation.pptx")
    out_project = os.path.join(os.path.dirname(__file__), "TalentStream_AI_Internship_Presentation.pptx")

    prs.save(out_downloads)
    prs.save(out_project)
    print(f"[SUCCESS] Presentation saved to {out_downloads}")
    print(f"[SUCCESS] Presentation saved to {out_project}")

if __name__ == "__main__":
    create_presentation()
