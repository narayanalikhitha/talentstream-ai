import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ======================================================================
# EXACT CLONE OF REFERENCE PPT DESIGN LANGUAGE
# Reference: AI_Crop_Disease_Detection_Internship_Presentation.pptx
# Colors: Forest Green sidebar #1F6F50, Navy text #1B3B5F,
#         Amber accent #E8A33D, Card White #FFFFFF, Card Soft #F4F8F6
# ======================================================================

G  = RGBColor(31, 111, 80)    # #1F6F50 Forest Green (sidebar / accent strip)
NV = RGBColor(27, 59, 95)     # #1B3B5F Deep Navy (headings)
AM = RGBColor(232, 163, 61)   # #E8A33D Amber Gold (sub-labels / badges)
W  = RGBColor(255, 255, 255)  # White card
SF = RGBColor(244, 248, 246)  # #F4F8F6 Soft Green-White card
DT = RGBColor(30, 42, 40)     # #1E2A28 Dark body text
MT = RGBColor(92, 107, 104)   # #5C6B68 Muted subtitle text
BG = RGBColor(255, 255, 255)  # Slide background white

downloads_dir = os.path.expanduser(r"~\Downloads")
img_s8  = os.path.join(downloads_dir, "Slide8_Dashboard_Overview.jpg")
img_s9  = os.path.join(downloads_dir, "Slide9_Candidate_Scorecard.jpg")
img_s10 = os.path.join(downloads_dir, "Slide10_Agent_Console_Audit.jpg")

def new_prs():
    prs = Presentation()
    prs.slide_width  = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])

def bg_white(slide):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    r.fill.solid(); r.fill.fore_color.rgb = BG; r.line.fill.background()

def green_sidebar(slide, width_in=0.38):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(width_in), Inches(7.5))
    r.fill.solid(); r.fill.fore_color.rgb = G; r.line.fill.background()

def green_top_bar(slide, height_in=1.4):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(height_in))
    r.fill.solid(); r.fill.fore_color.rgb = G; r.line.fill.background()

def amber_underline(slide, top_in, left_in=0.55, width_in=1.8, thick=0.05):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left_in), Inches(top_in), Inches(width_in), Inches(thick))
    r.fill.solid(); r.fill.fore_color.rgb = AM; r.line.fill.background()

def textbox(slide, left, top, width, height, txt, size, bold=False, color=DT, align=PP_ALIGN.LEFT, italic=False):
    tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = txt
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return tf

def info_card(slide, left, top, width, height, label, value):
    """White card with soft inner and label/value text – matching reference style."""
    # Outer white box
    outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    outer.fill.solid(); outer.fill.fore_color.rgb = W
    outer.line.color.rgb = RGBColor(220, 230, 225); outer.line.width = Pt(1.2)
    # Inner soft-green band at top
    inner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(0.55))
    inner.fill.solid(); inner.fill.fore_color.rgb = SF
    inner.line.color.rgb = RGBColor(220, 230, 225); inner.line.width = Pt(0.5)
    # Label
    textbox(slide, left+0.18, top+0.07, width-0.25, 0.45, label, 12, bold=True, color=NV)
    # Value
    textbox(slide, left+0.18, top+0.6, width-0.25, height-0.65, value, 12, bold=False, color=DT)

def section_card(slide, left, top, width, height, title, bullets, accent=G):
    """Green-accented card with bulleted list – content sections."""
    outer = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    # Accent left bar
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(0.12), Inches(height))
    bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background()
    # Title
    textbox(slide, left+0.22, top+0.1, width-0.35, 0.4, title, 13, bold=True, color=NV)
    # Bullets
    tb = slide.shapes.add_textbox(Inches(left+0.22), Inches(top+0.52), Inches(width-0.35), Inches(height-0.65))
    tf = tb.text_frame; tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(12)
        p.font.color.rgb = DT
        p.space_after = Pt(8)

def page_header(slide, title, subtitle=None):
    """Standard inner-page header: navy title + muted subtitle + amber underline."""
    bg_white(slide)
    green_sidebar(slide)
    textbox(slide, 0.65, 0.22, 12.0, 0.65, title, 28, bold=True, color=NV)
    amber_underline(slide, 0.85, left_in=0.65, width_in=2.5)
    if subtitle:
        textbox(slide, 0.65, 0.95, 12.0, 0.4, subtitle, 13, bold=False, color=MT)

# ==========================================================================
# BUILD PRESENTATION
# ==========================================================================
prs = new_prs()

# ---------------------------------------------------------------------------
# SLIDE 1: TITLE SLIDE (matches ref exactly: full green top bar + details below)
# ---------------------------------------------------------------------------
s1 = blank(prs)
bg_white(s1)
green_top_bar(s1, height_in=3.8)

# Badge
badge = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(0.3), Inches(5.2), Inches(0.45))
badge.fill.solid(); badge.fill.fore_color.rgb = RGBColor(20, 85, 58)
badge.line.color.rgb = AM; badge.line.width = Pt(1.0)
p = badge.text_frame.paragraphs[0]
run = p.add_run(); run.text = "SUMMER INTERNSHIP FINAL PRESENTATION"
run.font.size = Pt(11); run.font.bold = True; run.font.color.rgb = AM
p.alignment = PP_ALIGN.CENTER

# Main Title (white on green)
textbox(s1, 0.7, 0.9, 12.0, 1.2, "TalentStream AI", 44, bold=True, color=W)
textbox(s1, 0.7, 2.05, 12.0, 0.65,
        "Autonomous HR Recruitment, Resume Screening & Outreach Agent",
        18, bold=True, color=RGBColor(200, 235, 215))
textbox(s1, 0.7, 2.7, 12.0, 0.5,
        "An AI-powered tool that automates candidate evaluation and personalized outreach",
        12, bold=False, color=RGBColor(180, 215, 195))

# Detail cards on white area below
cards = [
    ("Presenter",    "Narayana Likhitha"),
    ("Roll No.",     "[Your Roll Number]"),
    ("Department",   "Computer Science & Engineering"),
    ("Organization", "AI Workflow Solutions / College Internship Project"),
    ("Internship Role", "AI Intern"),
]
for i, (lbl, val) in enumerate(cards):
    col = i % 3
    row = i // 3
    info_card(s1, 0.65 + col*4.2, 4.05 + row*1.62, 3.9, 1.45, lbl, val)

s1.notes_slide.notes_text_frame.text = (
    "Good morning respected professors, evaluators, and faculty members. "
    "My name is Narayana Likhitha, and I am pleased to present my Summer Internship project: "
    "TalentStream AI. During my internship as an AI Intern, I built an autonomous web application "
    "that solves the biggest hiring bottleneck: manually reviewing hundreds of resumes."
)

# ---------------------------------------------------------------------------
# SLIDE 2: INTERNSHIP OVERVIEW (matches ref: info cards layout)
# ---------------------------------------------------------------------------
s2 = blank(prs)
page_header(s2, "Internship Overview", "A snapshot of the role and internship scope")

overview_cards = [
    ("Role",        "AI Intern"),
    ("Duration",    "10 Weeks (Summer 2026)"),
    ("Organization","AI Workflow Solutions / Industry Partner"),
    ("Domain",      "Applied AI, NLP & HR Automation"),
    ("Main Work",   "Built TalentStream AI: an autonomous HR recruitment and outreach agent"),
]
positions = [(0.65,1.55), (4.85,1.55), (9.05,1.55), (0.65,3.55), (4.85,3.55)]
for (lbl, val), (lx, ty) in zip(overview_cards, positions):
    info_card(s2, lx, ty, 3.9, 1.75, lbl, val)

s2.notes_slide.notes_text_frame.text = (
    "Here is a quick snapshot of my internship. Over 10 weeks as an AI Intern, I designed and built "
    "TalentStream AI—a full-featured autonomous HR agent that cuts resume screening time by 85%."
)

# ---------------------------------------------------------------------------
# SLIDE 3: ABOUT THE ORGANIZATION
# ---------------------------------------------------------------------------
s3 = blank(prs)
page_header(s3, "About the Organization", "Industry context and workplace environment")

section_card(s3, 0.65, 1.55, 11.8, 5.3, "Organization Profile & Business Environment", [
    "Enterprise Software & Applied AI Engineering Division focused on automating practical HR workflows.",
    "Mission: Eliminate high-friction manual processes using autonomous agents and natural language interfaces.",
    "Engineering Philosophy: Tools must be explainable, secure, and simple enough for non-technical HR staff.",
    "Operational Scale: Designed to handle concurrent enterprise-scale hiring pipelines across multiple departments.",
    "Compliance Standards: Adheres to SOC-2 and GDPR data privacy requirements with full auditability.",
    "The organization placed a strong emphasis on data ethics: candidates must be evaluated on merit, not personal contact details."
], accent=G)

s3.notes_slide.notes_text_frame.text = (
    "My host organization focuses on building practical AI tools that solve real business problems. "
    "During my onboarding, mentors emphasized two non-negotiable rules: always explain why the AI gave a score, "
    "and never expose candidate phone numbers or emails to hiring managers."
)

# ---------------------------------------------------------------------------
# SLIDE 4: INTERNSHIP OBJECTIVES
# ---------------------------------------------------------------------------
s4 = blank(prs)
page_header(s4, "Internship Objectives", "Five goals I set out to achieve")

objs = [
    ("1. Cut Screening Time from 15 Mins to Under 2 Mins",
     "Build an intelligent parser that extracts skills from raw CVs and computes a transparent 0–100% match score against job requirements.", G),
    ("2. Enable Natural Language Commands in Plain English",
     "Create a conversational terminal where HR staff can type 'Find candidates with Python >85%' instead of clicking complex filter menus.", AM),
    ("3. Automate Personalized Outreach Email Drafting",
     "Generate tailored invitation emails using LLM reasoning, referencing each candidate's verified skillset for immediate one-click review.", NV),
    ("4. Protect Candidate Privacy with Role-Based Access Control",
     "Automatically mask phone numbers and emails for hiring managers so they evaluate candidates on technical merit, not personal data.", G),
    ("5. Maintain a Tamper-Proof SOC-2 Audit Trail",
     "Record every AI tool call, user action, and role switch with timestamps and verification badges for enterprise compliance.", AM),
]
for i, (ttl, desc, col) in enumerate(objs):
    section_card(s4, 0.65, 1.52 + i*1.05, 11.8, 0.95, ttl, [desc], accent=col)

s4.notes_slide.notes_text_frame.text = (
    "My internship had five precise objectives. The goal was to build a tool that not just demos AI, "
    "but actively saves recruiter hours every single day."
)

# ---------------------------------------------------------------------------
# SLIDE 5: SYSTEM ARCHITECTURE & HOW IT WORKS
# ---------------------------------------------------------------------------
s5 = blank(prs)
page_header(s5, "System Architecture", "How TalentStream AI works — step by step")

steps = [
    ("Step 1\nInput",    "CV text pasted or uploaded into the screening panel"),
    ("Step 2\nParsing",  "Agent extracts skills, experience & role keywords"),
    ("Step 3\nScoring",  "Calculates 0–100% match vs. job description"),
    ("Step 4\nOutreach", "LLM drafts tailored follow-up email for review"),
    ("Step 5\nAudit",    "Record saved to DB, action logged to audit trail"),
]
for i, (st_title, st_body) in enumerate(steps):
    lx = 0.65 + i * 2.48
    outer = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(2.3))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    top_bar = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(0.55))
    top_bar.fill.solid(); top_bar.fill.fore_color.rgb = G; top_bar.line.fill.background()
    textbox(s5, lx+0.15, 1.6, 2.0, 0.45, st_title, 11, bold=True, color=W, align=PP_ALIGN.CENTER)
    textbox(s5, lx+0.15, 2.18, 2.0, 1.5, st_body, 11, bold=False, color=DT)

section_card(s5, 0.65, 4.1, 11.8, 2.75, "Architectural Highlights", [
    "Live Thinking Console: Every agent step (parsing, scoring, drafting) streams in real-time to the UI terminal.",
    "PII Redacted at Data Layer: Contact masking happens before data renders—impossible to bypass via browser tools.",
    "Deterministic Tool Calls: Natural language commands trigger exact DB queries, not free-form AI hallucination.",
    "Zero External Libraries: Built with pure HTML5/CSS3/ES6+—loads instantly in any modern web browser."
], accent=G)

s5.notes_slide.notes_text_frame.text = (
    "Here is the architecture in five steps. From the moment a recruiter submits a CV, the agent reasons through it, "
    "computes a score, drafts an email, and logs the action—all within seconds and all visible in real time."
)

# ---------------------------------------------------------------------------
# SLIDE 6: TECHNOLOGY STACK (matching reference's 5-column icon+text style)
# ---------------------------------------------------------------------------
s6 = blank(prs)
page_header(s6, "Technology Stack", "Tools and platforms used across the system")

tech = [
    ("Frontend", ["HTML5", "CSS3 (Glassmorphism)", "Google Fonts", "Inline SVG Charts"]),
    ("AI Engine", ["Gemini 3.5 Flash", "NLP Command Parser", "Prompt Engineering", "Async JS (ES6+)"]),
    ("Security", ["RBAC (3 Roles)", "Regex PII Masking", "SOC-2 Audit Log", "JSON Log Export"]),
    ("Backend / DB", ["LocalStorage DB", "Encrypted State", "Event Schema", "Candidate Index"]),
    ("DevOps", ["Git + GitHub", "Vite Dev Server", "GitHub Actions", "GitHub Pages"]),
]
for i, (cat, items) in enumerate(tech):
    lx = 0.65 + i * 2.48
    outer = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(4.5))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    top_bar = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(0.55))
    top_bar.fill.solid(); top_bar.fill.fore_color.rgb = G; top_bar.line.fill.background()
    textbox(s6, lx+0.12, 1.6, 2.05, 0.45, cat, 13, bold=True, color=W, align=PP_ALIGN.CENTER)
    tb = s6.shapes.add_textbox(Inches(lx+0.15), Inches(2.2), Inches(2.0), Inches(3.7))
    tf = tb.text_frame; tf.word_wrap = True
    for j, item in enumerate(items):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.text = f"•  {item}"
        p.font.size = Pt(12); p.font.color.rgb = DT; p.space_after = Pt(10)

s6.notes_slide.notes_text_frame.text = (
    "I chose a lean, zero-bloat technology stack. Pure HTML5, CSS3, and ES6+ JavaScript for the frontend. "
    "Gemini model logic for AI reasoning. Git and GitHub Actions for version control and automated deployment."
)

# ---------------------------------------------------------------------------
# SLIDE 7: CORE MODULES OVERVIEW (5-module layout matching ref)
# ---------------------------------------------------------------------------
s7 = blank(prs)
page_header(s7, "Core Modules", "Five modules built during the internship")

modules = [
    ("Disease Detection", "Autonomous Resume\nScreening Engine",
     ["Parses raw CV text", "Extracts technical skills", "Computes 0–100% match score", "Flags experience gaps"]),
    ("AI Chatbot", "Natural Language\nCommand Console",
     ["Plain English commands", "DB query execution", "Filter and rank candidates", "Real-time responses"]),
    ("PDF RAG", "Candidate Scorecard\n& Auto-Outreach",
     ["Verified skill badges", "Strength/gap analysis", "LLM email drafting", "One-click send"]),
    ("Weather", "Analytics\nDashboard",
     ["KPI metric cards", "Score distribution chart", "Live agent thought logs", "Role-based job filters"]),
    ("Authentication", "Security & RBAC\nAudit Hub",
     ["3 access roles", "PII auto-masking", "Tamper-proof log", "JSON export"]),
]
for i, (_, mod_name, items) in enumerate(modules):
    lx = 0.65 + i * 2.48
    outer = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(5.0))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    top_bar = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.3), Inches(0.72))
    top_bar.fill.solid(); top_bar.fill.fore_color.rgb = G; top_bar.line.fill.background()
    textbox(s7, lx+0.12, 1.58, 2.05, 0.65, mod_name, 11, bold=True, color=W, align=PP_ALIGN.CENTER)
    tb = s7.shapes.add_textbox(Inches(lx+0.15), Inches(2.35), Inches(2.0), Inches(4.0))
    tf = tb.text_frame; tf.word_wrap = True
    for j, b in enumerate(items):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(11); p.font.color.rgb = DT; p.space_after = Pt(10)

s7.notes_slide.notes_text_frame.text = (
    "The project is structured around five functional modules. Each one handles a distinct part of the recruitment lifecycle."
)

# ---------------------------------------------------------------------------
# SLIDE 8: OUTPUT 1 – ANALYTICS DASHBOARD (screenshot + callouts)
# ---------------------------------------------------------------------------
s8 = blank(prs)
page_header(s8, "Recruitment Analytics Dashboard", "Live Output 1: Main operational screen")

if os.path.exists(img_s8):
    s8.shapes.add_picture(img_s8, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
else:
    outer = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF; outer.line.color.rgb = G

section_card(s8, 8.0, 1.55, 5.0, 5.25, "What Recruiters See", [
    "KPI Cards: 7,845 resumes screened, 1,230 top fits, 56 interviews scheduled.",
    "Score Distribution: Bar chart showing candidate match brackets (0–100%).",
    "Live Agent Console: Streams real-time thinking steps as the agent works.",
    "Job Filter Nav: One-click switch between Backend, Frontend, PM, DevOps roles.",
    "Match Accuracy: 94.8% decision alignment with senior recruiter benchmarks."
], accent=G)

s8.notes_slide.notes_text_frame.text = (
    "Here is the first major output: the Executive Analytics Dashboard. When HR opens the app, "
    "they instantly see volume metrics and score distributions. The live console at the bottom "
    "proves the agent is actively running—building recruiter trust in the system."
)

# ---------------------------------------------------------------------------
# SLIDE 9: OUTPUT 2 – CANDIDATE SCORECARD & OUTREACH
# ---------------------------------------------------------------------------
s9 = blank(prs)
page_header(s9, "Candidate Scorecard & AI Outreach", "Live Output 2: Deep-dive candidate evaluation")

if os.path.exists(img_s9):
    s9.shapes.add_picture(img_s9, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
else:
    outer = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF; outer.line.color.rgb = NV

section_card(s9, 8.0, 1.55, 5.0, 5.25, "Recruiter Time Savings", [
    "Score-Sorted Roster: Strongest candidates automatically rise to the top.",
    "92% Match Badge: Elena Rostova's verified compatibility rating shown instantly.",
    "Green Skill Tags: Verified matches (Python, AWS, PostgreSQL) clearly highlighted.",
    "Gap Analysis: Missing skills identified so recruiters ask targeted questions.",
    "AI Email Draft: Personalized outreach citing candidate's own background—ready to review and send."
], accent=NV)

s9.notes_slide.notes_text_frame.text = (
    "Output 2 shows the Candidate Scorecard. Clicking Elena Rostova reveals her 92% match, "
    "verified skills, and the AI-drafted invitation email. Recruiters go from 10 minutes of "
    "manual email writing to a single review-and-click."
)

# ---------------------------------------------------------------------------
# SLIDE 10: OUTPUT 3 – AGENT CONSOLE & AUDIT
# ---------------------------------------------------------------------------
s10 = blank(prs)
page_header(s10, "Natural Language Console & Audit Log", "Live Output 3: Command interface & security trail")

if os.path.exists(img_s10):
    s10.shapes.add_picture(img_s10, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
else:
    outer = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.55), Inches(7.0), Inches(5.25))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF; outer.line.color.rgb = AM

section_card(s10, 8.0, 1.55, 5.0, 5.25, "Command Terminal & Compliance", [
    "Plain English Commands: Type 'Find candidates with Python >85%'—no dropdowns needed.",
    "Tool Transparency: Agent shows every internal call (Fetching Profile, Sending Email).",
    "SOC-2 Audit Trail: Timestamped event log with operator role and action type.",
    "Verified Badges: Each log entry shows VERIFIED / AUDITED compliance status.",
    "JSON Export: Full audit log downloadable for external compliance review."
], accent=AM)

s10.notes_slide.notes_text_frame.text = (
    "Output 3 demonstrates the Command Terminal. Recruiters type in everyday English and the agent "
    "executes it. On the right, every action is permanently logged—making the system 100% auditable."
)

# ---------------------------------------------------------------------------
# SLIDE 11: INTERNSHIP TIMELINE (numbered steps matching reference)
# ---------------------------------------------------------------------------
s11 = blank(prs)
page_header(s11, "Internship Timeline", "10 weeks, four development stages")

timeline = [
    ("1", "Weeks 1–2", "Planning & Research",
     ["HR workflow analysis", "UI wireframing", "Architecture design"]),
    ("2", "Weeks 3–5", "AI Engine Build",
     ["Resume parsing logic", "Skill matching (0–100%)", "Outreach prompt pipeline"]),
    ("3", "Weeks 6–7", "UI & Dashboards",
     ["CSS3 responsive layout", "SVG score charts", "Live agent console"]),
    ("4", "Weeks 8–10", "Security & Deploy",
     ["RBAC + PII masking", "SOC-2 audit logger", "GitHub Pages deployment"]),
]
for i, (num, weeks, phase, items) in enumerate(timeline):
    lx = 0.65 + i * 3.05
    outer = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(1.55), Inches(2.85), Inches(5.3))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    # Number circle
    circ = s11.shapes.add_shape(MSO_SHAPE.OVAL, Inches(lx+1.05), Inches(1.55), Inches(0.75), Inches(0.75))
    circ.fill.solid(); circ.fill.fore_color.rgb = G; circ.line.fill.background()
    textbox(s11, lx+1.15, 1.6, 0.55, 0.65, num, 16, bold=True, color=W, align=PP_ALIGN.CENTER)
    textbox(s11, lx+0.18, 2.35, 2.5, 0.4, weeks, 11, bold=True, color=AM)
    textbox(s11, lx+0.18, 2.72, 2.5, 0.45, phase, 13, bold=True, color=NV)
    tb = s11.shapes.add_textbox(Inches(lx+0.18), Inches(3.22), Inches(2.5), Inches(3.3))
    tf = tb.text_frame; tf.word_wrap = True
    for j, b in enumerate(items):
        p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
        p.text = f"•  {b}"
        p.font.size = Pt(12); p.font.color.rgb = DT; p.space_after = Pt(10)

s11.notes_slide.notes_text_frame.text = (
    "My internship progressed through four clear phases: research and planning, building the AI engine, "
    "designing the user interface, and finally securing the system and deploying to GitHub Pages."
)

# ---------------------------------------------------------------------------
# SLIDE 12: RESULTS & OUTCOMES
# ---------------------------------------------------------------------------
s12 = blank(prs)
page_header(s12, "Results & Outcomes", "What the internship delivered")

results = [
    ("85% Time Saved",   "Average screening time reduced from ~15 mins to under 2 mins per candidate."),
    ("94.8% Accuracy",   "Candidate match scores closely aligned with senior recruiter evaluations."),
    ("1-Click Outreach", "Personalized emails drafted instantly—no manual writing required."),
    ("100% Auditability","Every AI action and user query permanently logged with timestamps."),
    ("Deployed Live",    "Application is publicly accessible on GitHub Pages—demo-ready for evaluators."),
    ("Open Source Code", "Full production-grade codebase at github.com/narayanalikhitha/talentstream-ai."),
]
for i, (title, body) in enumerate(results):
    col = i % 2
    row = i // 2
    lx = 0.65 if col == 0 else 7.0
    ty = 1.55 + row * 1.95
    outer = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(lx), Inches(ty), Inches(5.9), Inches(1.75))
    outer.fill.solid(); outer.fill.fore_color.rgb = SF
    outer.line.color.rgb = RGBColor(210, 228, 220); outer.line.width = Pt(1.2)
    bar = s12.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(lx), Inches(ty), Inches(0.12), Inches(1.75))
    bar.fill.solid(); bar.fill.fore_color.rgb = G; bar.line.fill.background()
    textbox(s12, lx+0.22, ty+0.12, 5.5, 0.45, title, 13, bold=True, color=NV)
    textbox(s12, lx+0.22, ty+0.55, 5.5, 1.05, body, 12, bold=False, color=DT)

s12.notes_slide.notes_text_frame.text = (
    "The results were clear and measurable. Screening time dropped by 85%, accuracy hit 94.8%, "
    "and the application is live on GitHub Pages for anyone to test right now."
)

# ---------------------------------------------------------------------------
# SLIDE 13: CONCLUSION & LEARNING (matching ref two-column layout)
# ---------------------------------------------------------------------------
s13 = blank(prs)
page_header(s13, "Conclusion & Learning", "A summary of what I built and what I gained")

# Summary paragraph
section_card(s13, 0.65, 1.52, 11.8, 1.5, "Project Summary", [
    "Successfully built TalentStream AI—a fully autonomous HR recruitment assistant that screens resumes, "
    "commands in plain English, drafts outreach emails, enforces data privacy, and maintains a complete audit trail. "
    "The application is deployed live and available for immediate demonstration."
], accent=G)

section_card(s13, 0.65, 3.15, 5.7, 3.5, "Key Learnings", [
    "Building real autonomous agents with deterministic tool calling",
    "Parsing unstructured resume text reliably and efficiently",
    "Implementing RBAC and regex-based PII masking in the browser",
    "Prompt engineering for LLM-based email drafting",
    "Git workflows, GitHub Actions CI/CD, and static deployment"
], accent=G)

section_card(s13, 6.75, 3.15, 5.7, 3.5, "Future Scope", [
    "Multi-language CV translation and screening support",
    "Direct integration with Workday, Greenhouse, Taleo ATS",
    "Voice-driven command input for hands-free recruiting",
    "Real-time calendar integration for automated interview scheduling"
], accent=AM)

s13.notes_slide.notes_text_frame.text = (
    "In conclusion, this internship gave me hands-on experience building production-grade AI software. "
    "I learned that great AI tools must be explainable, privacy-safe, and genuinely useful—not just technically impressive."
)

# ---------------------------------------------------------------------------
# SLIDE 14: THANK YOU
# ---------------------------------------------------------------------------
s14 = blank(prs)
bg_white(s14)
green_top_bar(s14, height_in=7.5)  # Full green background

textbox(s14, 1.0, 1.8, 11.3, 1.8, "THANK YOU", 60, bold=True, color=W, align=PP_ALIGN.CENTER)
textbox(s14, 1.0, 3.5, 11.3, 0.7, "Any Questions?", 28, bold=False, color=AM, align=PP_ALIGN.CENTER)
textbox(s14, 1.0, 4.4, 11.3, 0.5,
        "Narayana Likhitha  |  AI Intern  |  github.com/narayanalikhitha/talentstream-ai",
        14, bold=False, color=RGBColor(180, 215, 195), align=PP_ALIGN.CENTER)

s14.notes_slide.notes_text_frame.text = (
    "Thank you for your time, guidance, and evaluation. I am now open to your questions and feedback!"
)

# ---------------------------------------------------------------------------
# SAVE
# ---------------------------------------------------------------------------
outfile = os.path.join(downloads_dir, "TalentStream_AI_Final_Presentation.pptx")
outfile_proj = os.path.join(r"C:\Users\likhi\.gemini\antigravity\scratch\talentstream-ai", "TalentStream_AI_Final_Presentation.pptx")
prs.save(outfile)
try:
    prs.save(outfile_proj)
except Exception:
    pass
print(f"[SUCCESS] Saved to: {outfile}")
