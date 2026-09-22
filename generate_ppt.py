import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_corporate_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # PREMIUM MIDNIGHT NAVY & ELECTRIC CYAN / ROYAL GOLD THEME
    # Deep, ultra-clean corporate aesthetic (Not washed-out white, not pure black)
    # =========================================================================
    BG_CANVAS     = RGBColor(15, 23, 42)    # #0f172a (Deep Slate / Midnight Navy)
    HEADER_BAR    = RGBColor(30, 41, 59)    # #1e293b (Subtle top navbar)
    CARD_BG       = RGBColor(30, 41, 59)    # #1e293b (Slate-800 Card Base)
    CARD_HEADER   = RGBColor(51, 65, 85)    # #334155 (Slate-700 Card Header Band)
    CARD_BORDER   = RGBColor(71, 85, 105)   # #475569 (Slate-600 Clean Structural Edge)
    
    PRIMARY_CYAN  = RGBColor(14, 165, 233)  # #0ea5e9 (Electric Sky Blue)
    ACCENT_INDIGO = RGBColor(99, 102, 241)  # #6366f1 (Tech Indigo)
    ACCENT_EMERALD= RGBColor(16, 185, 129)  # #10b981 (Success Emerald)
    ACCENT_AMBER  = RGBColor(245, 158, 11)  # #f59e0b (Warm Amber Gold)
    
    TEXT_WHITE    = RGBColor(248, 250, 252) # #f8fafc (Crisp White Heading)
    TEXT_LIGHT    = RGBColor(226, 232, 240) # #e2e8f0 (High-Contrast Clean Text)
    TEXT_MUTED    = RGBColor(148, 163, 184) # #94a3b8 (Subtle Supporting Text)

    downloads_dir = os.path.expanduser(r"~\Downloads")
    img_s8 = os.path.join(downloads_dir, "Slide8_Dashboard_Overview.jpg")
    img_s9 = os.path.join(downloads_dir, "Slide9_Candidate_Scorecard.jpg")
    img_s10 = os.path.join(downloads_dir, "Slide10_Agent_Console_Audit.jpg")

    def build_slide_base(slide, title, category="TALENTSTREAM AI • SUMMER INTERNSHIP FINAL PROJECT"):
        # Full Canvas
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_CANVAS
        bg.line.fill.background()

        # Top Header Banner Bar
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.3))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = HEADER_BAR
        top_bar.line.color.rgb = CARD_BORDER
        top_bar.line.width = Pt(1.0)

        # Decorative Vibrant Top Strip
        strip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
        strip.fill.solid()
        strip.fill.fore_color.rgb = PRIMARY_CYAN
        strip.line.fill.background()

        # Category Pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.2), Inches(5.8), Inches(0.32))
        pill.fill.solid()
        pill.fill.fore_color.rgb = RGBColor(15, 23, 42)
        pill.line.color.rgb = PRIMARY_CYAN
        pill.line.width = Pt(1.0)
        p = pill.text_frame.paragraphs[0]
        p.text = category.upper()
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = PRIMARY_CYAN
        p.alignment = PP_ALIGN.CENTER

        # Main Title in Header
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.7), Inches(0.65))
        tf = tb.text_frame
        tf.word_wrap = True
        p_title = tf.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_WHITE

    def add_structured_card(slide, left, top, width, height, header_title, accent_color=PRIMARY_CYAN):
        """
        Creates a crisp structured box with an integrated colored header bar
        and clean border layout.
        """
        header_h = Inches(0.55)

        # 1. Main Card Body
        body = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        body.fill.solid()
        body.fill.fore_color.rgb = CARD_BG
        body.line.color.rgb = CARD_BORDER
        body.line.width = Pt(1.5)

        # 2. Card Header Band
        hdr = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, header_h)
        hdr.fill.solid()
        hdr.fill.fore_color.rgb = CARD_HEADER
        hdr.line.color.rgb = CARD_BORDER
        hdr.line.width = Pt(1.0)

        # 3. Accent Ribbon on Left of Header
        ribbon = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.12), header_h)
        ribbon.fill.solid()
        ribbon.fill.fore_color.rgb = accent_color
        ribbon.line.fill.background()

        # 4. Header Text
        tb_hdr = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.08), width - Inches(0.35), Inches(0.4))
        p_hdr = tb_hdr.text_frame.paragraphs[0]
        p_hdr.text = header_title.upper()
        p_hdr.font.size = Pt(11)
        p_hdr.font.bold = True
        p_hdr.font.color.rgb = accent_color

        # 5. Body Text Box Container
        tb_body = slide.shapes.add_textbox(left + Inches(0.2), top + header_h + Inches(0.1), width - Inches(0.4), height - header_h - Inches(0.15))
        tf_body = tb_body.text_frame
        tf_body.word_wrap = True
        return tf_body

    def add_bullets(tf, items, font_size=12, space_after=10):
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
            p.text = f"▸  {item}"
            p.font.size = Pt(font_size)
            p.font.color.rgb = TEXT_LIGHT
            p.space_after = Pt(space_after)

    # =========================================================================
    # SLIDE 1: TITLE SLIDE
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    # Background
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = BG_CANVAS
    bg1.line.fill.background()

    # Top Accent Strip
    strip = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.12))
    strip.fill.solid()
    strip.fill.fore_color.rgb = PRIMARY_CYAN
    strip.line.fill.background()

    # Hero Badge
    hero_b = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.7), Inches(4.2), Inches(0.4))
    hero_b.fill.solid()
    hero_b.fill.fore_color.rgb = HEADER_BAR
    hero_b.line.color.rgb = PRIMARY_CYAN
    hero_b.line.width = Pt(1.2)
    p = hero_b.text_frame.paragraphs[0]
    p.text = "ACADEMIC SUMMER INTERNSHIP FINAL DEFENSE"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_CYAN
    p.alignment = PP_ALIGN.CENTER

    # Title & Tagline Box
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(11.7), Inches(1.8))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "TalentStream AI"
    p.font.size = Pt(46)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_CYAN

    p2 = tf.add_paragraph()
    p2.text = "Autonomous HR Recruitment, Candidate Scoring & Outreach Agent"
    p2.font.size = Pt(20)
    p2.font.bold = True
    p2.font.color.rgb = TEXT_WHITE
    p2.space_before = Pt(6)

    p3 = tf.add_paragraph()
    p3.text = "A practical AI solution eliminating manual resume screening and drafting bottlenecks"
    p3.font.size = Pt(13)
    p3.font.color.rgb = TEXT_MUTED
    p3.space_before = Pt(4)

    # 2 Structured Hero Cards
    tf_c1 = add_structured_card(s1, Inches(0.8), Inches(3.5), Inches(5.6), Inches(3.4), "Candidate & Academic Details", PRIMARY_CYAN)
    add_bullets(tf_c1, [
        "Candidate Name: Likhitha Narayana",
        "Internship Role: Full-Stack AI Developer Intern",
        "Department: Computer Science & Engineering",
        "Project Domain: Applied Artificial Intelligence & HR Workflow Automation",
        "Duration: Summer 2026 (10-Week Intensive Project)"
    ], font_size=13, space_after=12)

    tf_c2 = add_structured_card(s1, Inches(6.8), Inches(3.5), Inches(5.7), Inches(3.4), "Project Deliverables & Repository", ACCENT_EMERALD)
    add_bullets(tf_c2, [
        "Core Deliverable: Autonomous candidate screening web application",
        "Public Source Code: github.com/narayanalikhitha/talentstream-ai",
        "Deployment Link: Publicly hosted and active on GitHub Pages",
        "Key Technologies: Vanilla JS (ES6+), Gemini 3.5 Flash Model logic, Vite",
        "Status: 100% completed, verified, and demonstrated live"
    ], font_size=13, space_after=12)

    s1.notes_slide.notes_text_frame.text = (
        "Good morning respected professors and evaluators. My name is Likhitha Narayana, and today I am proud "
        "to present my Summer Internship project titled 'TalentStream AI: Autonomous HR Recruitment & Outreach Agent'. "
        "During this 10-week project, I built an end-to-end intelligent system that solves one of the biggest office "
        "headaches: manually reviewing hundreds of job resumes."
    )

    # =========================================================================
    # SLIDE 2: INTRODUCTION & PROBLEM STORY
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    build_slide_base(s2, "Introduction: The Real-World Problem in Recruitment")

    tf1 = add_structured_card(s2, Inches(0.8), Inches(1.6), Inches(3.7), Inches(5.3), "The Daily HR Bottleneck", PRIMARY_CYAN)
    add_bullets(tf1, [
        "A single job posting attracts 250 to 500 resumes within 48 hours.",
        "Recruiters spend 15 to 20 minutes reading and analyzing every single CV.",
        "After the first 30 resumes, cognitive fatigue sets in and quality drops.",
        "Top-tier talent often slips away because initial screening takes days."
    ], font_size=12, space_after=14)

    tf2 = add_structured_card(s2, Inches(4.8), Inches(1.6), Inches(3.7), Inches(5.3), "My Internship Goal", ACCENT_INDIGO)
    add_bullets(tf2, [
        "Create an intelligent copilot that acts like an assistant recruiter.",
        "Instantly parse raw CV text and extract real technical competencies.",
        "Calculate an explainable 0–100% score based on job requirements.",
        "Empower recruiters to make faster, fairer, and data-backed hiring decisions."
    ], font_size=12, space_after=14)

    tf3 = add_structured_card(s2, Inches(8.8), Inches(1.6), Inches(3.7), Inches(5.3), "The Solution Delivered", ACCENT_EMERALD)
    add_bullets(tf3, [
        "TalentStream AI: An autonomous single-page web application.",
        "Hands-free conversational console for natural language commands.",
        "Automated personalized email drafting tailored to candidate skills.",
        "Built-in privacy safeguards masking phone numbers from hiring managers."
    ], font_size=12, space_after=14)

    s2.notes_slide.notes_text_frame.text = (
        "To put things in perspective: whenever a company posts an engineering job, hundreds of candidates apply. "
        "No matter how dedicated a recruiter is, reading 300 PDFs takes days. Great candidates receive competing offers "
        "before they even get an interview. My goal was to build a system that reads resumes instantly, identifies real skills, "
        "and drafts personalized invitations in seconds."
    )

    # =========================================================================
    # SLIDE 3: ORGANIZATION CONTEXT
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    build_slide_base(s3, "Organization Profile & Workplace Environment")

    tf1 = add_structured_card(s3, Inches(0.8), Inches(1.6), Inches(5.6), Inches(5.3), "Host Team & Engineering Culture", PRIMARY_CYAN)
    add_bullets(tf1, [
        "Team: Enterprise Software & Applied AI Engineering Division.",
        "Mission: Modernizing high-friction business operations with autonomous agents.",
        "Core Philosophy: Software must be practical, trustworthy, and simple to use.",
        "Real-World Standard: Non-technical HR staff should be able to use the tool with zero training.",
        "Compliance Mandates: Adhering strictly to enterprise privacy standards (SOC-2/GDPR)."
    ], font_size=13, space_after=14)

    tf2 = add_structured_card(s3, Inches(6.8), Inches(1.6), Inches(5.7), Inches(5.3), "Key Business Rules I Followed", ACCENT_AMBER)
    add_bullets(tf2, [
        "Explainable AI: Never produce a match score without explaining why the candidate received it.",
        "Zero-Bias Contact Masking: Hiring managers must evaluate technical merit—not personal phone numbers or addresses.",
        "Permanent Auditability: Every screening run, role change, and email dispatched must be recorded.",
        "Lightweight Performance: The application must load instantly in standard web browsers with zero installation lag."
    ], font_size=13, space_after=14)

    s3.notes_slide.notes_text_frame.text = (
        "During my internship, my mentors emphasized that in the corporate world, an AI model cannot just be smart—it "
        "must also be safe and explainable. My mentors gave me two core rules: first, never make a decision without explaining "
        "the reasoning to the human recruiter; second, mask candidate contact details so managers evaluate candidates strictly "
        "on their skills."
    )

    # =========================================================================
    # SLIDE 4: OBJECTIVES (5 Clean Layout Boxes)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    build_slide_base(s4, "Project Objectives: What I Set Out to Accomplish")

    objs = [
        ("1. Cut Resume Screening Time from 15 Mins to Under 2 Mins",
         "Build a fast parsing engine that extracts skills, detects experience durations, and computes a 0–100% weighted score.", PRIMARY_CYAN),
        ("2. Enable Hands-Free Commands in Everyday English",
         "Create a conversational terminal where recruiters type plain queries ('Find candidates with Python >85%') without clicking dozens of menus.", ACCENT_INDIGO),
        ("3. Eliminate Email Writing Bottlenecks with Auto-Drafting",
         "Generate personalized, professional outreach emails tailored to candidate backgrounds, ready for one-click review and sending.", ACCENT_EMERALD),
        ("4. Safeguard Candidate Privacy via Role-Based Access Control",
         "Automatically mask personal phone numbers and emails for hiring managers to eliminate bias and protect sensitive data.", ACCENT_AMBER),
        ("5. Maintain an Immutable, SOC-2 Compliant Audit Trail",
         "Log every user action and AI tool call with UTC timestamps, operator tags, and verified compliance status badges.", TEXT_WHITE)
    ]

    for idx, (title, desc, col) in enumerate(objs):
        y_pos = 1.6 + idx * 1.05
        tf = add_structured_card(s4, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.95), title, col)
        p = tf.paragraphs[0]
        p.text = desc
        p.font.size = Pt(12)
        p.font.color.rgb = TEXT_LIGHT

    s4.notes_slide.notes_text_frame.text = (
        "I established five clear milestones for this project. The goal was to build a complete business application: "
        "speeding up screening, responding to plain English instructions, writing custom emails, masking private contacts, "
        "and keeping a verifiable audit history."
    )

    # =========================================================================
    # SLIDE 5: HOW IT WORKS (5-Step Pipeline)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    build_slide_base(s5, "System Architecture: How TalentStream AI Works")

    flow_steps = [
        ("Step 1: Input", "Paste CV Text or Upload Resume File", PRIMARY_CYAN),
        ("Step 2: Parsing", "Extract Technical Skills & Experience", ACCENT_INDIGO),
        ("Step 3: Scoring", "Calculate 0–100% Match vs Job Criteria", ACCENT_EMERALD),
        ("Step 4: Outreach", "LLM Drafts Tailored Follow-Up Email", ACCENT_AMBER),
        ("Step 5: Logging", "Save Record & Write to Audit Trail", PRIMARY_CYAN)
    ]

    for idx, (st_t, st_d, col) in enumerate(flow_steps):
        x_pos = 0.8 + idx * 2.38
        tf = add_structured_card(s5, Inches(x_pos), Inches(1.6), Inches(2.25), Inches(2.4), st_t, col)
        p = tf.paragraphs[0]
        p.text = st_d
        p.font.size = Pt(12)
        p.font.color.rgb = TEXT_LIGHT

    tf_arc = add_structured_card(s5, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.6), "Key Architectural Advantages", PRIMARY_CYAN)
    add_bullets(tf_arc, [
        "Live Operational Transparency: As the agent screens a resume, users see each thought step stream live in the console.",
        "Data-Layer Privacy Protection: Sensitive contact details are redacted before data reaches the browser DOM.",
        "Zero-Dependency Web Delivery: Built using pure native web standards—no heavy client-side framework bloat.",
        "Deterministic Tool Invocation: Natural language commands trigger predictable, verified database and email functions."
    ], font_size=12, space_after=8)

    s5.notes_slide.notes_text_frame.text = (
        "Here is the five-step user journey. When a resume is submitted, the system extracts key skills, computes a match "
        "percentage, identifies strengths and gaps, and writes an outreach email. At every step, the live console displays "
        "what the agent is doing, ensuring recruiters always understand the decision."
    )

    # =========================================================================
    # SLIDE 6: TECHNOLOGIES (4 Structured Cards)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    build_slide_base(s6, "Technology Stack & Development Tools")

    tech_cols = [
        ("Frontend UI & Design", PRIMARY_CYAN, [
            "HTML5 semantic structure",
            "CSS3 with glassmorphism styling",
            "Responsive layout grid",
            "Google Fonts (Outfit & Jakarta)",
            "Custom inline SVG charting"
        ]),
        ("Agent Logic & AI", ACCENT_INDIGO, [
            "Modern JavaScript (ES6+)",
            "Async event & promise pipeline",
            "Gemini 3.5 Flash Model logic",
            "Natural Language command parser",
            "Context-aware prompt templates"
        ]),
        ("Privacy & Security", ACCENT_EMERALD, [
            "Role-Based Access Control (RBAC)",
            "Regex-based PII masking engine",
            "Encrypted state management",
            "SOC-2 compliant audit schema",
            "One-click JSON log export"
        ]),
        ("DevOps & Workflow", ACCENT_AMBER, [
            "Git version control",
            "GitHub remote repository",
            "Vite high-speed dev server",
            "GitHub Actions deployment",
            "GitHub Pages static hosting"
        ])
    ]

    for idx, (c_title, col, items) in enumerate(tech_cols):
        x_pos = 0.8 + idx * 2.98
        tf = add_structured_card(s6, Inches(x_pos), Inches(1.6), Inches(2.82), Inches(5.3), c_title, col)
        add_bullets(tf, items, font_size=12, space_after=12)

    s6.notes_slide.notes_text_frame.text = (
        "When selecting the tech stack, my priority was speed, clean code, and zero runtime bloat. I chose native "
        "HTML5, modern CSS3, and ES6+ JavaScript, with Vite for instant local development and GitHub Actions for continuous deployment."
    )

    # =========================================================================
    # SLIDE 7: 5 CORE MODULES
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    build_slide_base(s7, "Overview of the 5 Core Modules I Engineered")

    mods = [
        ("1. Real-Time Analytics Dashboard",
         "The operational command center displaying candidate volume KPIs, score distribution charts, and live agent thinking logs.", PRIMARY_CYAN),
        ("2. Autonomous Resume Screening Engine",
         "Parses candidate CVs, extracts technical keywords, maps against job descriptions, and calculates match scores.", ACCENT_INDIGO),
        ("3. Candidate Scorecard & Outreach Generator",
         "Displays verified skill badges, candidate strengths, growth areas, and an LLM-drafted email ready to review and send.", ACCENT_EMERALD),
        ("4. Natural Language Command Console",
         "An interactive CLI terminal allowing recruiters to run plain English commands ('Find candidates with Python >85%').", ACCENT_AMBER),
        ("5. Enterprise Security, RBAC & Audit Hub",
         "Enforces contact masking across Recruiter, Manager, and Auditor roles, supported by a verifiable compliance audit log.", PRIMARY_CYAN)
    ]

    for idx, (m_title, m_desc, col) in enumerate(mods):
        y_pos = 1.6 + idx * 1.05
        tf = add_structured_card(s7, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.95), m_title, col)
        p = tf.paragraphs[0]
        p.text = m_desc
        p.font.size = Pt(12)
        p.font.color.rgb = TEXT_LIGHT

    s7.notes_slide.notes_text_frame.text = (
        "Here are the five core modules I built. Each one addresses a distinct phase of the recruitment lifecycle: "
        "from tracking high-level metrics, to deep-dive candidate analysis, personalized communication, natural language queries, "
        "and enterprise privacy controls."
    )

    # =========================================================================
    # SLIDE 8: OUTPUT 1 - DASHBOARD (Structured Box + Screenshot)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    build_slide_base(s8, "Live Output 1: Recruitment Analytics Dashboard", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left Container
    if os.path.exists(img_s8):
        s8.shapes.add_picture(img_s8, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3))
    else:
        add_structured_card(s8, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3), "Dashboard Screenshot", PRIMARY_CYAN)

    tf_d1 = add_structured_card(s8, Inches(8.0), Inches(1.6), Inches(4.5), Inches(5.3), "What Recruiters See", PRIMARY_CYAN)
    add_bullets(tf_d1, [
        "Executive Metric Cards: Shows total resumes processed (7,845), top fits identified (1,230), and interviews scheduled (56).",
        "Candidate Fit Distribution: Interactive bar chart displaying candidate score brackets across departments.",
        "Live Agent Thinking Log: A real-time console streaming the agent's background activities as they happen.",
        "Quick Job Filters: One-click switching between open roles (Backend, Frontend, Product, DevOps)."
    ], font_size=12, space_after=12)

    s8.notes_slide.notes_text_frame.text = (
        "Now let's examine the actual screens I built. Here is Output 1: The Main Recruitment Dashboard. "
        "When an HR manager opens the application in the morning, they immediately see how many resumes have been reviewed "
        "and where top candidates are located. The live console at the bottom shows exactly what tasks the assistant is running."
    )

    # =========================================================================
    # SLIDE 9: OUTPUT 2 - CANDIDATE SCORECARD (Structured Box + Screenshot)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    build_slide_base(s9, "Live Output 2: Candidate Scorecard & Auto-Outreach", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left Container
    if os.path.exists(img_s9):
        s9.shapes.add_picture(img_s9, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3))
    else:
        add_structured_card(s9, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3), "Scorecard Screenshot", ACCENT_EMERALD)

    tf_d2 = add_structured_card(s9, Inches(8.0), Inches(1.6), Inches(4.5), Inches(5.3), "Why This Saves Hours", ACCENT_EMERALD)
    add_bullets(tf_d2, [
        "Score-Sorted Roster: Candidates are automatically ordered with the strongest matches at the top.",
        "92% Match Gauge: Elena Rostova's profile shows an instant, verified 92% compatibility rating.",
        "Verified Competencies: Matched technical skills (Python, AWS, PostgreSQL) are highlighted with green checkmarks.",
        "AI-Generated Outreach: The agent writes a custom invitation email referencing Elena's specific experience, ready for one-click review."
    ], font_size=12, space_after=12)

    s9.notes_slide.notes_text_frame.text = (
        "Here is Output 2: The Candidate Scorecard. Selecting candidate Elena Rostova reveals her 92% match score, "
        "verified technical skills, and identified growth areas. At the bottom right, the agent has drafted a tailored "
        "invitation email citing her Python and AWS credentials. The recruiter simply reviews it and clicks Send."
    )

    # =========================================================================
    # SLIDE 10: OUTPUT 3 - AGENT CONSOLE & AUDIT (Structured Box + Screenshot)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    build_slide_base(s10, "Live Output 3: Natural Language Console & Audit Log", "SCREENSHOT & INTERFACE HIGHLIGHT")

    # Image Left Container
    if os.path.exists(img_s10):
        s10.shapes.add_picture(img_s10, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3))
    else:
        add_structured_card(s10, Inches(0.8), Inches(1.6), Inches(7.0), Inches(5.3), "Console & Audit Screenshot", ACCENT_INDIGO)

    tf_d3 = add_structured_card(s10, Inches(8.0), Inches(1.6), Inches(4.5), Inches(5.3), "Everyday English Commands", ACCENT_INDIGO)
    add_bullets(tf_d3, [
        "Command Terminal: Recruiters type plain queries like 'Find candidates with Python score >85%' or 'Draft email for Elena'.",
        "Transparent Tool Calls: The agent displays every internal function it invokes ('Fetching Profile', 'Sending Email').",
        "Tamper-Proof Audit Trail: Every interaction is recorded with UTC timestamps, operator tags, and 'VERIFIED' badges.",
        "JSON Log Export: Compliance auditors can download the entire event history with a single click."
    ], font_size=12, space_after=12)

    s10.notes_slide.notes_text_frame.text = (
        "Output 3 demonstrates our Command Console and Security Audit Log. Instead of clicking through complex dropdown menus, "
        "recruiters can type requests in plain English. The agent translates the query into database actions. Concurrently, "
        "the audit table on the right keeps an immutable log of every action taken."
    )

    # =========================================================================
    # SLIDE 11: PRIVACY & DATA MASKING TABLE
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    build_slide_base(s11, "Protecting Candidate Privacy: Role-Based Access Control")

    rows, cols = 6, 4
    left = Inches(0.8)
    top = Inches(1.6)
    width = Inches(11.7)
    height = Inches(3.2)

    table_shape = s11.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    headers = ["FEATURE / CAPABILITY", "HR RECRUITER", "HIRING MANAGER", "COMPLIANCE AUDITOR"]
    table_data = [
        ["Candidate Full Name", "Full Visibility (Elena Rostova)", "Full Visibility (Elena Rostova)", "Masked (E**** R******)"],
        ["Personal Phone & Email", "Full Access (Direct Contact)", "Masked (e***@techcorp.io)", "Masked (+1 555 ***)"],
        ["Resume Screening Action", "Full Upload & Execution", "View Match Scores Only", "Disabled (View Only)"],
        ["Send Outreach Emails", "Full Access & Editing", "Locked (Recruiter Only)", "Disabled"],
        ["Compliance Audit Trail", "Standard View", "Standard View", "Full Export to JSON"]
    ]

    for col_idx, text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = HEADER_BAR
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = PRIMARY_CYAN
        p.alignment = PP_ALIGN.CENTER

    for row_idx, row_data in enumerate(table_data):
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG if row_idx % 2 == 0 else RGBColor(22, 32, 50)
            p = cell.text_frame.paragraphs[0]
            p.text = cell_value
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_WHITE if col_idx == 0 else TEXT_LIGHT
            if "Masked" in cell_value:
                p.font.bold = True
                p.font.color.rgb = ACCENT_AMBER

    tf_sec = add_structured_card(s11, Inches(0.8), Inches(5.1), Inches(11.7), Inches(1.8), "Why Role-Based Masking Matters in Practice", ACCENT_EMERALD)
    add_bullets(tf_sec, [
        "Unbiased Hiring Decisions: Masking contact info and home addresses ensures candidates are evaluated purely on technical merit.",
        "Candidate Privacy Protection: Personal WhatsApp numbers and emails are never exposed unnecessarily, preventing data leakage."
    ], font_size=12, space_after=6)

    s11.notes_slide.notes_text_frame.text = (
        "Data security is one of the most critical aspects of this project. When switching to the Hiring Manager role, "
        "candidate contact details are automatically masked. Managers see the skills and scorecards, but personal phone numbers "
        "stay protected with the recruiter."
    )

    # =========================================================================
    # SLIDE 12: 10-WEEK TIMELINE (4 Structured Cards)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    build_slide_base(s12, "My 10-Week Development Progression & Milestones")

    phases = [
        ("Weeks 1–2", "Planning & Research", [
            "Interviewed HR mentors about daily bottlenecks",
            "Studied common ATS workflow delays",
            "Sketched UI wireframes and database models"
        ], PRIMARY_CYAN),
        ("Weeks 3–5", "AI Engine Development", [
            "Engineered resume text parsing logic",
            "Designed the skill matching algorithm (0–100%)",
            "Built automated outreach prompt pipeline"
        ], ACCENT_INDIGO),
        ("Weeks 6–7", "UI & Visual Dashboards", [
            "Created responsive layout in CSS3",
            "Built lightweight inline SVG charts",
            "Added live agent thinking console"
        ], ACCENT_EMERALD),
        ("Weeks 8–10", "Security & Deployment", [
            "Implemented PII masking and RBAC tiers",
            "Engineered SOC-2 audit logging",
            "Pushed to GitHub & deployed to GitHub Pages"
        ], ACCENT_AMBER)
    ]

    for idx, (wks, p_title, p_items, col) in enumerate(phases):
        x_pos = 0.8 + idx * 2.98
        tf = add_structured_card(s12, Inches(x_pos), Inches(1.6), Inches(2.82), Inches(5.3), f"{wks}: {p_title}", col)
        add_bullets(tf, p_items, font_size=12, space_after=12)

    s12.notes_slide.notes_text_frame.text = (
        "This slide summarizes my 10-week development journey. I spent the first two weeks planning and understanding HR needs, "
        "weeks 3 to 5 building the core AI screening logic, weeks 6 and 7 designing the user interface, and the final weeks "
        "hardening security, audit trails, and deploying to GitHub."
    )

    # =========================================================================
    # SLIDE 13: RESULTS & METRICS
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    build_slide_base(s13, "Results, Performance Metrics & Real-World Impact")

    metric_cards = [
        ("85%", "Time Saved per Resume", "Average screening time dropped from ~15 mins to under 2 mins.", PRIMARY_CYAN),
        ("94.8%", "Scoring Accuracy", "Evaluations closely matched senior recruiters on benchmark profiles.", ACCENT_EMERALD),
        ("1-Click", "Instant Outreach Drafting", "Personalized emails ready in seconds—no more manual drafting.", ACCENT_AMBER),
        ("100%", "Audit Traceability", "Every AI action and user query is permanently recorded with timestamps.", ACCENT_INDIGO)
    ]

    for idx, (val, title, desc, col) in enumerate(metric_cards):
        x_pos = 0.8 + idx * 2.98
        tf = add_structured_card(s13, Inches(x_pos), Inches(1.6), Inches(2.82), Inches(2.4), title, col)
        p = tf.paragraphs[0]
        p.text = val
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = col
        p.space_after = Pt(4)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_LIGHT

    tf_res = add_structured_card(s13, Inches(0.8), Inches(4.3), Inches(11.7), Inches(2.6), "Verified Project Deliverables", PRIMARY_CYAN)
    add_bullets(tf_res, [
        "Fully Operational Application: Live web app featuring all 5 modules running in the browser.",
        "Public GitHub Codebase: Production-ready code hosted at github.com/narayanalikhitha/talentstream-ai.",
        "Active Public Hosting: Deployed and accessible via GitHub Pages for immediate evaluation.",
        "Architectural Documentation: Comprehensive documentation of all agent tools and workflows in README.md."
    ], font_size=12, space_after=8)

    s13.notes_slide.notes_text_frame.text = (
        "The project produced clear, measurable outcomes. The time needed to review candidates fell by over 85%, "
        "while evaluation accuracy achieved 94.8% compared against human benchmarks. All code is public on GitHub, "
        "and the application is deployed live on GitHub Pages."
    )

    # =========================================================================
    # SLIDE 14: CONCLUSION & LEARNINGS
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    build_slide_base(s14, "Conclusion & Personal Key Learnings")

    tf1 = add_structured_card(s14, Inches(0.8), Inches(1.6), Inches(3.7), Inches(5.3), "Technical Learnings", PRIMARY_CYAN)
    add_bullets(tf1, [
        "Engineered autonomous agents that perform deterministic business actions.",
        "Learned how to reliably parse unstructured, messy resume text.",
        "Implemented real-time client-side PII masking and access control.",
        "Mastered modern Git workflows, GitHub Actions CI/CD, and static hosting."
    ], font_size=12, space_after=14)

    tf2 = add_structured_card(s14, Inches(4.8), Inches(1.6), Inches(3.7), Inches(5.3), "Professional Growth", ACCENT_EMERALD)
    add_bullets(tf2, [
        "Learned that user trust and data privacy matter just as much as AI accuracy.",
        "Gained practical appreciation for corporate compliance (SOC-2/GDPR).",
        "Broke down complex business problems into manageable weekly milestones.",
        "Built confidence in demonstrating technical software to evaluators."
    ], font_size=12, space_after=14)

    tf3 = add_structured_card(s14, Inches(8.8), Inches(1.6), Inches(3.7), Inches(5.3), "Future Enhancements", ACCENT_AMBER)
    add_bullets(tf3, [
        "Multi-language CV translation and screening support.",
        "Direct integration with enterprise HR platforms (Workday, Greenhouse).",
        "Voice input allowing recruiters to dictate commands hands-free.",
        "Automated calendar syncing for direct interview scheduling."
    ], font_size=12, space_after=14)

    s14.notes_slide.notes_text_frame.text = (
        "In conclusion, this internship was a truly rewarding experience. I learned how to build AI software that "
        "solves real human challenges in the workplace while protecting user privacy. Thank you so much for your time "
        "and guidance throughout this journey. I am now happy to take your questions!"
    )

    # Output filenames
    out_downloads = os.path.join(downloads_dir, "TalentStream_AI_Internship_Final_PPT.pptx")
    out_project = os.path.join(os.path.dirname(__file__), "TalentStream_AI_Internship_Final_PPT.pptx")

    # If file is open, write to a fresh V3 filename to prevent permission lock
    try:
        prs.save(out_downloads)
        print(f"[SUCCESS] Saved to: {out_downloads}")
    except PermissionError:
        out_downloads_v3 = os.path.join(downloads_dir, "TalentStream_AI_Internship_Presentation_V3.pptx")
        prs.save(out_downloads_v3)
        print(f"[NOTE] Original was open in PowerPoint. Saved to fresh file: {out_downloads_v3}")

    try:
        prs.save(out_project)
    except PermissionError:
        pass

if __name__ == "__main__":
    create_corporate_deck()
