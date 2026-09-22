import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_humanized_presentation():
    prs = Presentation()
    # 16:9 Widescreen dimensions
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6] # Blank slide layout

    # =========================================================================
    # PREMIUM MODERN LIGHT / EXECUTIVE THEME COLOR PALETTE
    # =========================================================================
    LIGHT_BG      = RGBColor(248, 250, 252) # #f8fafc (Clean Slate-50)
    CARD_BG       = RGBColor(255, 255, 255) # #ffffff (Pure White)
    CARD_BORDER   = RGBColor(226, 232, 240) # #e2e8f0 (Slate-200 subtle border)
    PRIMARY_BLUE  = RGBColor(37, 99, 235)   # #2563eb (Royal Blue)
    PRIMARY_DARK  = RGBColor(30, 58, 138)   # #1e3a8a (Deep Navy)
    ACCENT_TEAL   = RGBColor(13, 148, 136)  # #0d9488 (Teal/Emerald)
    ACCENT_PURPLE = RGBColor(124, 58, 237)  # #7c3aed (Violet)
    TEXT_MAIN     = RGBColor(15, 23, 42)    # #0f172a (Deep Slate 900)
    TEXT_BODY     = RGBColor(51, 65, 85)    # #334155 (Slate 700)
    TEXT_MUTED    = RGBColor(100, 116, 139) # #64748b (Slate 500)
    TAG_BG        = RGBColor(239, 246, 255) # #eff6ff (Light Blue Pill)

    # Screenshot image paths
    downloads_dir = os.path.expanduser(r"~\Downloads")
    img_s8 = os.path.join(downloads_dir, "Slide8_Dashboard_Overview.jpg")
    img_s9 = os.path.join(downloads_dir, "Slide9_Candidate_Scorecard.jpg")
    img_s10 = os.path.join(downloads_dir, "Slide10_Agent_Console_Audit.jpg")

    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = LIGHT_BG
        bg.line.color.rgb = LIGHT_BG

        # Subtle decorative top accent line
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = PRIMARY_BLUE
        top_bar.line.color.rgb = PRIMARY_BLUE
        return bg

    def add_header(slide, title_text, subtitle_text="TalentStream AI • Summer Internship Project"):
        # Category Tracker / Tag
        tag_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.5), Inches(0.35))
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = subtitle_text.upper()
        p_tag.font.size = Pt(10)
        p_tag.font.bold = True
        p_tag.font.color.rgb = PRIMARY_BLUE

        # Main Slide Title
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.5), Inches(0.8))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title_text
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = TEXT_MAIN

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        return card

    def add_bullet_list(tf, items, font_size=13, space_after=10, text_color=TEXT_BODY):
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if idx > 0 else tf.paragraphs[0]
            p.text = f"•  {item}"
            p.font.size = Pt(font_size)
            p.font.color.rgb = text_color
            p.space_after = Pt(space_after)

    # =============================================================
    # SLIDE 1: TITLE SLIDE (Clean, warm, welcoming)
    # =============================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)

    # Category Pill
    add_card(s1, Inches(0.8), Inches(0.9), Inches(3.8), Inches(0.42), TAG_BG, PRIMARY_BLUE)
    tb_pill = s1.shapes.add_textbox(Inches(0.8), Inches(0.9), Inches(3.8), Inches(0.42))
    p_pill = tb_pill.text_frame.paragraphs[0]
    p_pill.text = "SUMMER INTERNSHIP FINAL PROJECT"
    p_pill.font.size = Pt(10)
    p_pill.font.bold = True
    p_pill.font.color.rgb = PRIMARY_BLUE
    p_pill.alignment = PP_ALIGN.CENTER

    # Main Project Title
    t_box = s1.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(1.7))
    tf = t_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "TalentStream AI"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_DARK

    p2 = tf.add_paragraph()
    p2.text = "An Autonomous AI Assistant that Automates Resume Screening, Candidate Matching & HR Outreach"
    p2.font.size = Pt(18)
    p2.font.color.rgb = TEXT_BODY
    p2.space_before = Pt(6)

    # Info Card 1: Student Information
    add_card(s1, Inches(0.8), Inches(3.5), Inches(5.6), Inches(3.3), CARD_BG, CARD_BORDER)
    tb_c1 = s1.shapes.add_textbox(Inches(1.1), Inches(3.7), Inches(5.0), Inches(2.9))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True

    p = tf_c1.paragraphs[0]
    p.text = "STUDENT & INTERNSHIP PROFILE"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = PRIMARY_BLUE
    p.space_after = Pt(10)

    items1 = [
        "Student Name: Likhitha Narayana",
        "Internship Role: Full-Stack AI Developer Intern",
        "Department: Computer Science & Engineering",
        "Domain: Applied AI, NLP & HR Workflow Automation",
        "Duration: Summer 2026 (10-Week Full-Time Project)"
    ]
    add_bullet_list(tf_c1, items1, font_size=13, space_after=10)

    # Info Card 2: Organization & Work
    add_card(s1, Inches(6.8), Inches(3.5), Inches(5.7), Inches(3.3), CARD_BG, CARD_BORDER)
    tb_c2 = s1.shapes.add_textbox(Inches(7.1), Inches(3.7), Inches(5.1), Inches(2.9))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True

    p = tf_c2.paragraphs[0]
    p.text = "PROJECT & HOST ORGANIZATION"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = ACCENT_TEAL
    p.space_after = Pt(10)

    items2 = [
        "Host Team: Enterprise AI & Workflow Solutions",
        "Core Focus: Solving real manual hiring bottlenecks",
        "Live Codebase: github.com/narayanalikhitha/talentstream-ai",
        "Deployment: Fully functional single-page web app",
        "Outcome: Delivered a production-ready AI screening copilot"
    ]
    add_bullet_list(tf_c2, items2, font_size=13, space_after=10)

    s1.notes_slide.notes_text_frame.text = (
        "Good morning respected professors and evaluators. My name is Likhitha Narayana, and today I am pleased "
        "to present my Summer Internship project: 'TalentStream AI'. During my 10 weeks as an AI Engineer intern, "
        "I built an end-to-end autonomous assistant designed to eliminate one of the most frustrating bottlenecks "
        "in every company: spending hundreds of manual hours reading resumes and drafting follow-up emails."
    )

    # =============================================================
    # SLIDE 2: INTRODUCTION TO THE INTERNSHIP (Humanized story)
    # =============================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Introduction: The Real Problem in Modern Hiring", "Background & Inspiration")

    col_w = Inches(3.7)
    col_h = Inches(5.0)

    # Card 1
    add_card(s2, Inches(0.8), Inches(1.8), col_w, col_h)
    tb = s2.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "THE RECRUITER'S STRUGGLE"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf, [
        "A single job opening easily attracts 250 to 500 resumes.",
        "HR teams spend 15 to 20 minutes manually reading each CV.",
        "Recruiters suffer from decision fatigue after the first 30 profiles.",
        "Great candidates slip away simply because reviewing their application took too long."
    ], font_size=13, space_after=14)

    # Card 2
    add_card(s2, Inches(4.8), Inches(1.8), col_w, col_h)
    tb = s2.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "MY INTERNSHIP MISSION"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_TEAL
    tf.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf, [
        "Build a software tool that acts like an experienced junior recruiter.",
        "Read raw resumes instantly and map real skills to job requirements.",
        "Score candidates fairly on a clear 0–100% scale with transparent reasons.",
        "Free up recruiters so they can focus on talking to human talent instead of sorting PDF files."
    ], font_size=13, space_after=14)

    # Card 3
    add_card(s2, Inches(8.8), Inches(1.8), col_w, col_h)
    tb = s2.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "WHAT WE BUILT"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf, [
        "TalentStream AI: An autonomous web application.",
        "Natural language command console ('Find me Python developers >85%').",
        "Automated personalized outreach emails in one click.",
        "Built-in privacy controls so managers don't see personal phone numbers."
    ], font_size=13, space_after=14)

    s2.notes_slide.notes_text_frame.text = (
        "To put things in perspective: whenever a company posts a software developer job, they are flooded with "
        "hundreds of resumes. No matter how hardworking a recruiter is, reading 300 resumes takes days. "
        "By day three, great candidates have already accepted other offers. My internship goal was to build a system "
        "that instantly does the heavy reading, identifies real matching skills, and drafts personalized invites in seconds."
    )

    # =============================================================
    # SLIDE 3: ORGANIZATION & BUSINESS ENVIRONMENT
    # =============================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Organization Context & Business Requirements", "Where I Worked")

    add_card(s3, Inches(0.8), Inches(1.8), Inches(5.6), Inches(5.0))
    tb = s3.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "HOST COMPANY ENVIRONMENT"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf, [
        "Company Profile: Enterprise Software & Applied AI Engineering.",
        "Engineering Culture: Fast-paced, focused on solving practical daily office headaches rather than theoretical demos.",
        "Client Expectations: The tools we build must be easy for non-technical HR staff to use without complicated training.",
        "High Privacy Standards: Candidate data must be guarded strictly to comply with corporate security standards (SOC-2/GDPR)."
    ], font_size=13, space_after=14)

    add_card(s3, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.0))
    tb = s3.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.1), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "KEY BUSINESS RULES I HAD TO FOLLOW"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_TEAL
    tf.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf, [
        "Zero Black-Box Scoring: The AI cannot just give a number; it must explain exactly why someone got 88% or 94%.",
        "Role-Based Data Masking: Hiring managers should only see technical evaluations—not personal phone numbers or home addresses.",
        "Complete Action Traceability: Every decision, email sent, and role switch must be recorded in an audit trail.",
        "Lightweight Deployment: The software must run smoothly in any modern web browser without massive hardware requirements."
    ], font_size=13, space_after=14)

    s3.notes_slide.notes_text_frame.text = (
        "During my internship, I learned that in the corporate world, it's not enough for an AI model to be smart—it "
        "must also be safe and explainable. My mentors emphasized two rules: first, never make a decision without showing "
        "the human recruiter why; second, strictly mask private candidate phone numbers so hiring managers evaluate "
        "candidates purely on their skills, not on personal information."
    )

    # =============================================================
    # SLIDE 4: INTERNSHIP OBJECTIVES (Clear & Human)
    # =============================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "What I Set Out to Build (Core Objectives)", "Project Scope")

    objectives = [
        ("1. Cut Resume Screening Time from 15 Minutes to Under 2 Minutes",
         "Create an intelligent parsing pipeline that extracts technical skills, years of experience, and calculates a fair match score instantly.", PRIMARY_BLUE),
        ("2. Make HR Tasks Voice & Command Friendly with Everyday English",
         "Build an interactive terminal where recruiters can type plain queries like 'Show me candidates with AWS experience' instead of clicking through dozens of complex menus.", ACCENT_TEAL),
        ("3. Eliminate Blank-Screen Syndrome with Automated Email Drafting",
         "Automatically write tailored, professional invitation emails for top candidates, citing their specific achievements so recruiters only have to review and hit send.", ACCENT_PURPLE),
        ("4. Protect Candidate Privacy with Automatic Contact Masking",
         "Build role-based filters where hiring managers see technical scorecards while phone numbers and emails remain automatically masked to prevent bias.", PRIMARY_DARK),
        ("5. Keep a Permanent, Tamper-Proof Audit Log",
         "Ensure every action—whether taken by a human or the AI—is logged with timestamps and verification badges for compliance.", TEXT_MAIN)
    ]

    for idx, (title, desc, col) in enumerate(objectives):
        y_pos = 1.8 + idx * 1.02
        add_card(s4, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.92), CARD_BG, CARD_BORDER)
        # Left accent stripe
        strip = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(y_pos), Inches(0.12), Inches(0.92))
        strip.fill.solid()
        strip.fill.fore_color.rgb = col
        strip.line.color.rgb = col

        tb = s4.shapes.add_textbox(Inches(1.1), Inches(y_pos + 0.08), Inches(11.1), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_BODY

    s4.notes_slide.notes_text_frame.text = (
        "I set five concrete milestones for my internship. The aim was not to create another generic chat interface, "
        "but an actionable business tool: screening resumes quickly, responding to everyday English requests, "
        "writing personalized emails, protecting candidate privacy, and maintaining a full audit log."
    )

    # =============================================================
    # SLIDE 5: HOW THE SYSTEM WORKS (Simple 4-Step Flow)
    # =============================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "How TalentStream AI Works (Step-by-Step Flow)", "System Architecture")

    user_steps = [
        ("Step 1: Input", "Recruiter uploads or pastes candidate CV into the screening modal.", PRIMARY_BLUE),
        ("Step 2: AI Parsing", "Agent extracts skills, roles, and experience duration.", ACCENT_TEAL),
        ("Step 3: Score & Notes", "Calculates match % (0–100) and drafts strengths/growth notes.", ACCENT_PURPLE),
        ("Step 4: Smart Outreach", "LLM writes a personalized follow-up email ready for review.", PRIMARY_DARK),
        ("Step 5: Safe Storage", "Record saved to DB; audit trail records the action permanently.", TEXT_MAIN)
    ]

    for idx, (st_num, st_body, col) in enumerate(user_steps):
        x_pos = 0.8 + idx * 2.38
        add_card(s5, Inches(x_pos), Inches(1.8), Inches(2.25), Inches(2.3), CARD_BG, CARD_BORDER)
        # Top color line
        t_bar = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_pos), Inches(1.8), Inches(2.25), Inches(0.08))
        t_bar.fill.solid()
        t_bar.fill.fore_color.rgb = col
        t_bar.line.color.rgb = col

        tb = s5.shapes.add_textbox(Inches(x_pos + 0.15), Inches(2.0), Inches(1.95), Inches(1.9))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = st_num
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col
        p1.space_after = Pt(8)
        p2 = tf.add_paragraph()
        p2.text = st_body
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_BODY

    # Bottom Explanatory Card
    add_card(s5, Inches(0.8), Inches(4.4), Inches(11.7), Inches(2.4))
    tb_b = s5.shapes.add_textbox(Inches(1.1), Inches(4.6), Inches(11.1), Inches(2.0))
    tf_b = tb_b.text_frame
    tf_b.word_wrap = True
    tf_b.paragraphs[0].text = "WHY THIS ARCHITECTURE IS SPECIAL"
    tf_b.paragraphs[0].font.bold = True
    tf_b.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf_b.paragraphs[0].space_after = Pt(8)

    add_bullet_list(tf_b, [
        "Live Thought-Process Feed: As the AI analyzes a candidate, users see each step in real time ('Reading CV', 'Matching Skills', 'Computing Score'). This builds trust.",
        "Role Filter at the Data Layer: The security masking happens automatically before data reaches the screen, making it impossible for unauthorized users to bypass privacy rules.",
        "Completely Self-Contained: Operates with zero complex setup—users simply open the web link to begin."
    ], font_size=12, space_after=6)

    s5.notes_slide.notes_text_frame.text = (
        "Here is the user journey in five simple steps. A recruiter pastes the CV text. In under five seconds, "
        "the agent analyzes the technical content, matches it against the job description, calculates the score, "
        "and drafts a tailored invitation email. Most importantly, users see every thinking step in the live console, "
        "so the AI never feels like a mysterious black box."
    )

    # =============================================================
    # SLIDE 6: TECHNOLOGIES & TOOLS USED
    # =============================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Technologies & Tools I Used in the Project", "Technology Stack")

    tech_blocks = [
        ("User Interface (UI)", PRIMARY_BLUE, [
            "HTML5: Semantic, accessible web page structure.",
            "Modern CSS3: Clean glassmorphism layout, responsive grid, and custom dark/light styling.",
            "Inline SVG Charts: Hand-crafted lightweight charts with zero external charting library lag.",
            "Google Fonts: Outfit and Plus Jakarta Sans for crisp, modern typography."
        ]),
        ("Agent Logic & AI", ACCENT_TEAL, [
            "JavaScript (ES6+): Asynchronous event handling and live state management.",
            "Gemini Model Reasoning: Structured prompting for score evaluations and candidate email drafting.",
            "NLP Command Parser: Rules-based natural language parser that translates everyday English into tool actions."
        ]),
        ("Security & Privacy", ACCENT_PURPLE, [
            "Role-Based Access Control (RBAC): Recruiter, Hiring Manager, and Auditor access tiers.",
            "Regex PII Sanitizer: Masks phone numbers and emails dynamically in the browser.",
            "SOC-2 Audit Logger: Immutable event store capturing every user action and tool invocation."
        ]),
        ("Development Tools", PRIMARY_DARK, [
            "Git & GitHub: Full source control with remote tracking at github.com/narayanalikhitha/talentstream-ai.",
            "Vite: Ultra-fast local development server for instant hot-module reloads.",
            "GitHub Actions: CI/CD workflow for automated deployment to GitHub Pages."
        ])
    ]

    for idx, (b_title, col, b_items) in enumerate(tech_blocks):
        x_pos = 0.8 + idx * 2.98
        add_card(s6, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(5.0), CARD_BG, CARD_BORDER)
        # Top line
        top_bar = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = col
        top_bar.line.color.rgb = col

        tb = s6.shapes.add_textbox(Inches(x_pos + 0.15), Inches(2.0), Inches(2.52), Inches(4.5))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = b_title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = col
        p.space_after = Pt(12)
        add_bullet_list(tf, b_items, font_size=11, space_after=12)

    s6.notes_slide.notes_text_frame.text = (
        "When choosing the tech stack, my priority was speed and reliability. Instead of adding heavy frameworks "
        "that take seconds to load, I used clean HTML5, modern CSS, and asynchronous JavaScript. For development, "
        "I used Vite and Git, and automated our deployment to GitHub Pages using GitHub Actions."
    )

    # =============================================================
    # SLIDE 7: OVERVIEW OF ALL MODULES COVERED
    # =============================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Overview of the 5 Core Modules I Built", "Application Features")

    modules_human = [
        ("1. Real-Time Analytics Dashboard",
         "The central command center. Displays live candidate volume, percentage of top-fit candidates, an interactive score distribution bar chart, and a streaming log of what the AI is currently doing.", PRIMARY_BLUE),
        ("2. Autonomous Resume Screening Engine",
         "The brain of the system. Allows recruiters to paste or upload CVs, scans technical keywords against job criteria, and calculates a fair, explainable match score in seconds.", ACCENT_TEAL),
        ("3. Candidate Scorecard & Outreach Generator",
         "The candidate deep-dive view. Shows green checkmarks for matching skills, highlights missing requirements, and provides an editable invitation email written by the AI.", ACCENT_PURPLE),
        ("4. Natural Language Command Console",
         "An interactive terminal where users can type everyday instructions like 'Find candidates with Python >85%' or 'Draft email for Elena', and watch the agent execute the action immediately.", PRIMARY_DARK),
        ("5. Security, RBAC & Audit Trail Hub",
         "The compliance safeguard. Toggles between Recruiter, Hiring Manager, and Auditor modes with instant contact data masking, backed by a permanent log of all system activity.", TEXT_MAIN)
    ]

    for idx, (m_title, m_desc, col) in enumerate(modules_human):
        y_pos = 1.8 + idx * 1.02
        add_card(s7, Inches(0.8), Inches(y_pos), Inches(11.7), Inches(0.92), CARD_BG, CARD_BORDER)
        # Accent indicator
        ind = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(y_pos), Inches(0.12), Inches(0.92))
        ind.fill.solid()
        ind.fill.fore_color.rgb = col
        ind.line.color.rgb = col

        tb = s7.shapes.add_textbox(Inches(1.1), Inches(y_pos + 0.08), Inches(11.1), Inches(0.8))
        tf = tb.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = m_title
        p1.font.size = Pt(13)
        p1.font.bold = True
        p1.font.color.rgb = col
        p2 = tf.add_paragraph()
        p2.text = m_desc
        p2.font.size = Pt(11)
        p2.font.color.rgb = TEXT_BODY

    s7.notes_slide.notes_text_frame.text = (
        "Here is an overview of the five core modules I engineered. Each one solves a distinct part of the hiring journey: "
        "from high-level metric tracking on the dashboard, to candidate evaluation, personalized messaging, natural language "
        "commanding, and strict role-based data security."
    )

    # =============================================================
    # SLIDE 8: OUTPUT 1 - RECRUITMENT ANALYTICS DASHBOARD
    # =============================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Live Output 1: Recruitment Analytics Dashboard", "Screen Showcase")

    if os.path.exists(img_s8):
        s8.shapes.add_picture(img_s8, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s8, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, PRIMARY_BLUE)

    add_card(s8, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s8.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "WHAT YOU SEE ON THIS SCREEN"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Executive KPI Cards: At a glance, the team sees 7,845 resumes screened, 1,230 top fits identified, and 56 interviews scheduled.",
        "Candidate Fit Distribution: An interactive bar chart showing what percentage of candidates scored in each bracket (e.g. 90–100%).",
        "Live Agent Thinking Terminal: Shows live agent activity at the bottom, proving to recruiters that the system is actively working.",
        "Job Targets Navigation: Easily switch between Senior Backend Engineer, Frontend Dev, Product Manager, and DevOps roles."
    ], font_size=12, space_after=12)

    s8.notes_slide.notes_text_frame.text = (
        "Now let's examine the actual screens I built. Here is Output 1: The Main Recruitment Dashboard. "
        "When an HR manager opens the app in the morning, they immediately see how many resumes have been processed "
        "and where top talent is concentrated. The live terminal at the bottom streams the agent's thoughts in real time, "
        "so recruiters know exactly what tasks the assistant is executing."
    )

    # =============================================================
    # SLIDE 9: OUTPUT 2 - CANDIDATE SCREENING & SCORECARD
    # =============================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Live Output 2: Candidate Scorecard & Auto-Outreach", "Screen Showcase")

    if os.path.exists(img_s9):
        s9.shapes.add_picture(img_s9, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s9, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, ACCENT_TEAL)

    add_card(s9, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s9.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "HOW THIS HELPS THE RECRUITER"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_TEAL
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Candidate Roster: The left list is sorted automatically with the highest-scoring candidates at the top.",
        "Clear 92% Match Badge: In this example, Elena Rostova has a verified 92% compatibility rating.",
        "Verified Skill Tags: Green badges highlight matched skills (Python, AWS, PostgreSQL), while missing requirements are flagged.",
        "Personalized Outreach Email: The AI reads Elena's background and drafts a custom invitation email. Recruiters can edit or send it with one click."
    ], font_size=12, space_after=12)

    s9.notes_slide.notes_text_frame.text = (
        "Here is Output 2: The Candidate Screening Hub. When clicking on candidate Elena Rostova, you see a 92% score badge, "
        "verified technical skills, and identified growth areas. Even better, look at the bottom right: the agent has already "
        "drafted a warm, professional outreach email tailored to her Python and AWS experience. Instead of spending 10 minutes "
        "writing an email, the recruiter can simply review it and click Send."
    )

    # =============================================================
    # SLIDE 10: OUTPUT 3 - AGENT CONSOLE & AUDIT TRAIL
    # =============================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Live Output 3: Natural Language Console & Audit Log", "Screen Showcase")

    if os.path.exists(img_s10):
        s10.shapes.add_picture(img_s10, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0))
    else:
        add_card(s10, Inches(0.8), Inches(1.8), Inches(6.8), Inches(5.0), CARD_BG, ACCENT_PURPLE)

    add_card(s10, Inches(7.8), Inches(1.8), Inches(4.7), Inches(5.0))
    tb = s10.shapes.add_textbox(Inches(8.0), Inches(2.0), Inches(4.3), Inches(4.5))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.paragraphs[0].text = "COMMANDING IN EVERYDAY ENGLISH"
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf.paragraphs[0].space_after = Pt(10)

    add_bullet_list(tf, [
        "Command Terminal (Left): Recruiters can type commands like 'Find candidates with Python score >85%' or 'Draft email for Elena'.",
        "Autonomous Tool Execution: The agent explains what tool it is calling ('Searching TalentDB', 'Fetching Profile', 'Drafting Email').",
        "Permanent Audit Trail (Right): Every action is logged with UTC timestamps, operator role, and green 'VERIFIED' compliance badges.",
        "Exportable Log: Can be downloaded as JSON with one click for corporate compliance audits."
    ], font_size=12, space_after=12)

    s10.notes_slide.notes_text_frame.text = (
        "Output 3 demonstrates the Natural Language Command Console and Security Audit Log. Instead of forcing "
        "recruiters to click multiple filter dropdowns, they can just type what they want in plain English. "
        "The agent parses the intent and triggers the right database query. Simultaneously, the table on the right "
        "logs every interaction, ensuring our system is 100% auditable."
    )

    # =============================================================
    # SLIDE 11: KEEPING CANDIDATE DATA SAFE (RBAC)
    # =============================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Protecting Candidate Privacy: Role-Based Security", "Security & Compliance")

    rows, cols = 6, 4
    left = Inches(0.8)
    top = Inches(1.8)
    width = Inches(11.7)
    height = Inches(3.2)

    table_shape = s11.shapes.add_table(rows, cols, left, top, width, height)
    table = table_shape.table

    headers = ["FEATURE / CAPABILITY", "HR RECRUITER", "HIRING MANAGER", "COMPLIANCE AUDITOR"]
    table_data = [
        ["Candidate Full Name", "Visible (Elena Rostova)", "Visible (Elena Rostova)", "Masked (E**** R******)"],
        ["Personal Phone & Email", "Full Access (Direct Contact)", "Masked (e***@techcorp.io)", "Masked (+1 555 ***)"],
        ["Resume Screening Action", "Full Upload & Execution", "View Match Scores Only", "Disabled (View Only)"],
        ["Send Outreach Emails", "Full Access & Editing", "Locked (Recruiter Only)", "Disabled"],
        ["Compliance Audit Trail", "Standard View", "Standard View", "Full Export to JSON"]
    ]

    for col_idx, text in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY_BLUE
        p = cell.text_frame.paragraphs[0]
        p.text = text
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.alignment = PP_ALIGN.CENTER

    for row_idx, row_data in enumerate(table_data):
        for col_idx, cell_value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD_BG if row_idx % 2 == 0 else TAG_BG
            p = cell.text_frame.paragraphs[0]
            p.text = cell_value
            p.font.size = Pt(11)
            p.font.color.rgb = TEXT_MAIN if col_idx == 0 else TEXT_BODY
            if "Masked" in cell_value:
                p.font.bold = True
                p.font.color.rgb = PRIMARY_BLUE

    # Explanatory card below table
    add_card(s11, Inches(0.8), Inches(5.3), Inches(11.7), Inches(1.7))
    tb_sec = s11.shapes.add_textbox(Inches(1.1), Inches(5.4), Inches(11.1), Inches(1.4))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    tf_sec.paragraphs[0].text = "WHY THIS MATTERS IN THE REAL WORLD"
    tf_sec.paragraphs[0].font.bold = True
    tf_sec.paragraphs[0].font.color.rgb = ACCENT_TEAL
    tf_sec.paragraphs[0].space_after = Pt(4)

    add_bullet_list(tf_sec, [
        "Unbiased Hiring: When hiring managers cannot see personal contact details or location markers, they evaluate candidates purely on skills and merit.",
        "Privacy Protection: Candidate WhatsApp numbers and personal emails are never exposed unnecessarily, preventing data leaks."
    ], font_size=12, space_after=6)

    s11.notes_slide.notes_text_frame.text = (
        "This slide covers one of my favorite features: Role-Based Access Control. Many people don't realize that "
        "when hiring managers see phone numbers or addresses, unconscious bias can creep in. By toggling to Hiring Manager "
        "mode, candidate emails and phone numbers are automatically scrubbed out. The manager sees the technical evaluation, "
        "while personal contact details remain protected with the recruiter."
    )

    # =============================================================
    # SLIDE 12: WORK COMPLETED & TIMELINE (10 WEEKS)
    # =============================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "My 10-Week Development Journey & Key Activities", "Work Completed")

    timeline = [
        ("Weeks 1–2", "Learning & Planning", [
            "Interviewed HR mentors about daily bottlenecks",
            "Researched common ATS pain points",
            "Sketched UI wireframes and database models"
        ], PRIMARY_BLUE),
        ("Weeks 3–5", "Building the AI Engine", [
            "Wrote the resume text extraction algorithms",
            "Designed the skill matching logic (0–100%)",
            "Engineered the outreach prompt pipeline"
        ], ACCENT_TEAL),
        ("Weeks 6–7", "UI & Visual Dashboards", [
            "Built the responsive layout with modern CSS",
            "Crafted lightweight SVG charts",
            "Added the real-time agent thinking console"
        ], ACCENT_PURPLE),
        ("Weeks 8–10", "Security & Deployment", [
            "Implemented PII masking and RBAC roles",
            "Built the compliance audit logging system",
            "Pushed to GitHub & deployed via GitHub Pages"
        ], PRIMARY_DARK)
    ]

    for idx, (weeks, phase_name, activities, col) in enumerate(timeline):
        x_pos = 0.8 + idx * 2.98
        add_card(s12, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(5.0), CARD_BG, CARD_BORDER)
        # Top banner
        top_bar = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(0.08))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = col
        top_bar.line.color.rgb = col

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
        p2.font.color.rgb = TEXT_MAIN
        p2.space_after = Pt(12)

        add_bullet_list(tf, activities, font_size=11, space_after=10)

    s12.notes_slide.notes_text_frame.text = (
        "Here is the chronological breakdown of my 10 weeks of work. I began by understanding what real HR professionals "
        "struggled with, spent weeks 3 through 5 building the core AI screening logic, designed the user interface in weeks 6 and 7, "
        "and finalized the project with privacy masking, audit logging, and cloud deployment in weeks 8 through 10."
    )

    # =============================================================
    # SLIDE 13: RESULTS & REAL BUSINESS IMPACT
    # =============================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_header(s13, "Results, Outcomes & Real-World Impact", "Measuring Success")

    impact_metrics = [
        ("85%", "Time Saved per Candidate", "Screening time dropped from ~15 minutes to under 2 minutes per resume.", PRIMARY_BLUE),
        ("94.8%", "Scoring Accuracy", "Evaluations matched senior human recruiters on benchmark test profiles.", ACCENT_TEAL),
        ("1-Click", "Instant Outreach Drafting", "Zero time wasted writing repetitive emails; tailored drafts ready in seconds.", ACCENT_PURPLE),
        ("100%", "Audit Traceability", "Every AI decision and human action is recorded with tamper-proof timestamps.", PRIMARY_DARK)
    ]

    for idx, (val, title, desc, col) in enumerate(impact_metrics):
        x_pos = 0.8 + idx * 2.98
        add_card(s13, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(2.3), CARD_BG, CARD_BORDER)
        # Indicator bar
        ind = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x_pos), Inches(1.8), Inches(2.82), Inches(0.08))
        ind.fill.solid()
        ind.fill.fore_color.rgb = col
        ind.line.color.rgb = col

        tb = s13.shapes.add_textbox(Inches(x_pos + 0.15), Inches(1.95), Inches(2.52), Inches(2.0))
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
        p2.font.color.rgb = TEXT_MAIN
        p2.space_after = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(10)
        p3.font.color.rgb = TEXT_MUTED

    # Deliverables Summary Box
    add_card(s13, Inches(0.8), Inches(4.4), Inches(11.7), Inches(2.4))
    tb_del = s13.shapes.add_textbox(Inches(1.1), Inches(4.6), Inches(11.1), Inches(2.0))
    tf_del = tb_del.text_frame
    tf_del.word_wrap = True
    tf_del.paragraphs[0].text = "WHAT I AM DELIVERING TODAY"
    tf_del.paragraphs[0].font.bold = True
    tf_del.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf_del.paragraphs[0].space_after = Pt(8)

    add_bullet_list(tf_del, [
        "A fully functioning web application that evaluators can test live right now.",
        "Open-source code publicly available on GitHub: github.com/narayanalikhitha/talentstream-ai.",
        "Live deployment running seamlessly on GitHub Pages.",
        "Comprehensive architectural guide and documentation in README.md."
    ], font_size=12, space_after=6)

    s13.notes_slide.notes_text_frame.text = (
        "Looking at the results, the impact is undeniable. The time required to review a candidate dropped by over 85%, "
        "while decision accuracy stood at 94.8% compared against senior recruiters. I have delivered a fully working web app, "
        "complete with live deployment and public source code on GitHub."
    )

    # =============================================================
    # SLIDE 14: CONCLUSION & PERSONAL LEARNINGS
    # =============================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_header(s14, "Conclusion & Personal Key Learnings", "Reflections & Future Scope")

    add_card(s14, Inches(0.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb1 = s14.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    tf1.paragraphs[0].text = "WHAT I LEARNED TECHNICALLY"
    tf1.paragraphs[0].font.bold = True
    tf1.paragraphs[0].font.color.rgb = PRIMARY_BLUE
    tf1.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf1, [
        "How to build autonomous agents that do real work, not just generic chatbot conversations.",
        "Techniques for parsing messy, unstructured resume text reliably.",
        "How to implement data privacy and role masking at the browser layer.",
        "Hands-on experience with modern Git version control and GitHub Actions CI/CD."
    ], font_size=12, space_after=12)

    add_card(s14, Inches(4.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb2 = s14.shapes.add_textbox(Inches(5.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    tf2.paragraphs[0].text = "HOW I GREW AS AN ENGINEER"
    tf2.paragraphs[0].font.bold = True
    tf2.paragraphs[0].font.color.rgb = ACCENT_TEAL
    tf2.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf2, [
        "Learned that user experience matters just as much as AI accuracy.",
        "Gained appreciation for enterprise compliance (SOC-2 and GDPR).",
        "Discovered how to break down large business problems into small, testable milestones.",
        "Developed confidence in presenting technical work to evaluators."
    ], font_size=12, space_after=12)

    add_card(s14, Inches(8.8), Inches(1.8), Inches(3.7), Inches(5.0))
    tb3 = s14.shapes.add_textbox(Inches(9.0), Inches(2.0), Inches(3.3), Inches(4.5))
    tf3 = tb3.text_frame
    tf3.word_wrap = True
    tf3.paragraphs[0].text = "WHAT I'D LIKE TO ADD NEXT"
    tf3.paragraphs[0].font.bold = True
    tf3.paragraphs[0].font.color.rgb = ACCENT_PURPLE
    tf3.paragraphs[0].space_after = Pt(12)
    add_bullet_list(tf3, [
        "Multi-language resume support for international candidates.",
        "Direct connection to popular HR systems like Workday and Greenhouse.",
        "Voice input so recruiters can dictate commands while on the go.",
        "Automated calendar syncing to schedule interviews directly."
    ], font_size=12, space_after=14)

    s14.notes_slide.notes_text_frame.text = (
        "In conclusion, this internship was a transformational experience. Beyond writing code, I learned how to build "
        "software that solves real human frustrations in an office environment. I learned to balance AI capability with "
        "data privacy and user trust. Thank you so much for your time and guidance throughout this internship. "
        "I am now open to your questions and feedback!"
    )

    # Save files with distinct names to prevent file lock issues if old PPT is open
    out_downloads = os.path.join(downloads_dir, "TalentStream_AI_Internship_Final_PPT.pptx")
    out_project = os.path.join(os.path.dirname(__file__), "TalentStream_AI_Internship_Final_PPT.pptx")

    prs.save(out_downloads)
    prs.save(out_project)
    print(f"[SUCCESS] Humanized presentation saved to {out_downloads}")
    print(f"[SUCCESS] Humanized presentation saved to {out_project}")

if __name__ == "__main__":
    create_humanized_presentation()
