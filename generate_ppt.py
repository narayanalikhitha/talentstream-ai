import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ======================================================================
# HUMAN-STYLE STUDENT PPT — BLUE THEME
# Roll No: 239x1a3391 | Dept: CSE-(AI&ML) | Name: Narayana Likhitha
# Color Palette: Deep Blue #184C8C, Navy #0F2C59, Amber #E8A33D
# ======================================================================

BL = RGBColor(24, 76, 140)     # Deep Royal Blue (replaced Green)
NV = RGBColor(15, 44, 89)      # Deep Navy Blue (headings)
AM = RGBColor(232, 163, 61)    # Amber Gold (accent/bullets/highlights)
W  = RGBColor(255, 255, 255)   # White
SF = RGBColor(240, 244, 248)   # Soft Blue-White card background
DT = RGBColor(30, 42, 54)      # Dark body text
MT = RGBColor(90, 105, 120)    # Muted subtitle text
LB = RGBColor(205, 224, 242)   # Light blue card border
DB = RGBColor(12, 45, 90)      # Badge dark blue

downloads_dir = os.path.expanduser(r"~\Downloads")
img_s8  = os.path.join(downloads_dir, "Slide8_Dashboard_Overview.jpg")
img_s9  = os.path.join(downloads_dir, "Slide9_Candidate_Scorecard.jpg")
img_s10 = os.path.join(downloads_dir, "Slide10_Agent_Console_Audit.jpg")

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

def blank():
    return prs.slides.add_slide(prs.slide_layouts[6])

def rect(slide, l, t, w, h, color, border=None, border_w=1.0):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    if border:
        sh.line.color.rgb = border; sh.line.width = Pt(border_w)
    else:
        sh.line.fill.background()
    return sh

def rrect(slide, l, t, w, h, color, border=None, border_w=1.2):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = color
    if border:
        sh.line.color.rgb = border; sh.line.width = Pt(border_w)
    else:
        sh.line.fill.background()
    return sh

def tb(slide, l, t, w, h, text, size, bold=False, color=DT, align=PP_ALIGN.LEFT, italic=False, wrap=True):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = wrap
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run()
    r.text = text; r.font.size = Pt(size)
    r.font.bold = bold; r.font.italic = italic
    r.font.color.rgb = color
    return tf

def add_bullets_tf(slide, l, t, w, h, items, size=12.5, color=DT, space=8):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = f"  {item}"
        p.font.size = Pt(size); p.font.color.rgb = color; p.space_after = Pt(space)
    return tf

def inner_page(slide, title, subtitle=None):
    """Standard inner slide: white bg, blue left bar, navy title, amber underline."""
    rect(slide, 0, 0, 13.333, 7.5, W)
    rect(slide, 0, 0, 0.35, 7.5, BL)
    tb(slide, 0.6, 0.2, 12.0, 0.7, title, 28, bold=True, color=NV)
    rect(slide, 0.6, 0.88, 2.2, 0.05, AM)
    if subtitle:
        tb(slide, 0.6, 0.96, 11.5, 0.4, subtitle, 13, color=MT)

# ==========================================================================
# SLIDE 1: TITLE SLIDE — Blue top banner, student info cards below
# ==========================================================================
s1 = blank()
rect(s1, 0, 0, 13.333, 7.5, W)
rect(s1, 0, 0, 13.333, 3.7, BL)

# Badge tag
rrect(s1, 0.7, 0.3, 5.5, 0.42, DB, AM, 1.0)
tb(s1, 0.7, 0.32, 5.5, 0.38, "SUMMER INTERNSHIP PRESENTATION", 10.5, bold=True, color=AM, align=PP_ALIGN.CENTER)

# Title
tb(s1, 0.7, 0.88, 12.0, 1.25, "TalentStream AI", 46, bold=True, color=W)
tb(s1, 0.7, 2.1, 11.5, 0.55,
   "Autonomous HR Recruitment, Candidate Screening & Outreach Agent",
   16, bold=True, color=RGBColor(200, 225, 255))
tb(s1, 0.7, 2.72, 11.0, 0.45,
   "AI-powered tool that eliminates manual resume screening and automates candidate outreach",
   12, color=RGBColor(180, 210, 245))

# Info boxes (5 student detail cards)
info = [
    ("Presenter",    "Narayana Likhitha"),
    ("Roll No.",     "239x1a3391"),
    ("Department",   "CSE-(AI&ML)"),
    ("Organization", "AI Workflow Solutions / Industry Partner"),
    ("Internship Role", "AI Intern"),
]
positions = [(0.7, 3.9), (4.0, 3.9), (7.3, 3.9), (0.7, 5.6), (4.0, 5.6)]
for (lbl, val), (lx, ty) in zip(info, positions):
    rrect(s1, lx, ty, 3.0, 1.52, W, LB, 1.0)
    rrect(s1, lx, ty, 3.0, 0.48, SF, LB, 0.8)
    tb(s1, lx+0.15, ty+0.06, 2.7, 0.38, lbl, 11, bold=True, color=NV)
    tb(s1, lx+0.15, ty+0.55, 2.7, 0.85, val, 12.5, bold=(lbl in ["Roll No.", "Department"]), color=DT)

s1.notes_slide.notes_text_frame.text = (
    "Good morning everyone. My name is Narayana Likhitha, Roll No. 239x1a3391 from CSE-(AI&ML). "
    "This is my Summer Internship presentation on TalentStream AI."
)

# ==========================================================================
# SLIDE 2: INTERNSHIP OVERVIEW — Blue themed info cards
# ==========================================================================
s2 = blank()
inner_page(s2, "Internship Overview", "A quick look at my role and what I did during the internship")

overview = [
    ("Role",         "AI Intern"),
    ("Duration",     "10 Weeks  (Summer 2026)"),
    ("Organization", "AI Workflow Solutions / Industry Partner"),
    ("Domain",       "Applied AI, NLP & HR Workflow Automation"),
    ("Main Work",    "Built TalentStream AI: an autonomous HR agent that screens resumes and automates outreach emails"),
]
positions2 = [(0.6, 1.5), (4.65, 1.5), (8.7, 1.5), (0.6, 3.85), (4.65, 3.85)]
for (lbl, val), (lx, ty) in zip(overview, positions2):
    rrect(s2, lx, ty, 3.75, 1.95, W, LB, 1.0)
    rrect(s2, lx, ty, 3.75, 0.52, SF, LB, 0.8)
    tb(s2, lx+0.15, ty+0.07, 3.45, 0.4, lbl, 12, bold=True, color=NV)
    tb(s2, lx+0.15, ty+0.58, 3.45, 1.25, val, 12.5, color=DT)

s2.notes_slide.notes_text_frame.text = (
    "Here is a quick snapshot of my internship. I worked as an AI Intern for 10 weeks and built "
    "TalentStream AI — a web application that screens resumes and automates outreach for recruiters."
)

# ==========================================================================
# SLIDE 3: ABOUT THE ORGANIZATION
# ==========================================================================
s3 = blank()
inner_page(s3, "About the Organization", "Context of where I worked and what they do")

rect(s3, 0.6, 1.42, 12.0, 0.02, BL)

points = [
    "Works in the field of Artificial Intelligence and enterprise software development.",
    "Mission is to replace manual, repetitive office workflows with smart automated tools.",
    "Engineering standards require all AI outputs to be explainable and human-verifiable.",
    "Non-technical HR staff must be able to use the tools with no special training.",
    "Strict privacy compliance: candidate contact details cannot be exposed to hiring managers.",
    "All system actions must be recorded in a permanent audit trail for compliance review.",
]
for i, pt in enumerate(points):
    ty = 1.55 + i * 0.95
    rrect(s3, 0.6, ty, 12.0, 0.82, SF, LB, 1.0)
    rect(s3, 0.6, ty, 0.1, 0.82, BL)
    tb(s3, 0.85, ty + 0.13, 11.5, 0.6, pt, 12.5, color=DT)

s3.notes_slide.notes_text_frame.text = (
    "My host organization focuses on building practical AI tools that solve real daily office problems. "
    "They placed a strong emphasis on explainability and data privacy from day one of my internship."
)

# ==========================================================================
# SLIDE 4: PROBLEM STATEMENT & OBJECTIVES
# ==========================================================================
s4 = blank()
inner_page(s4, "Problem Statement & Internship Objectives", "What problem exists and what I was asked to solve")

# Left: problem
rrect(s4, 0.6, 1.42, 5.8, 5.5, SF, LB, 1.0)
rect(s4, 0.6, 1.42, 0.1, 5.5, BL)
tb(s4, 0.85, 1.5, 5.35, 0.42, "The Problem", 14, bold=True, color=NV)
rect(s4, 0.85, 1.9, 1.8, 0.04, AM)
add_bullets_tf(s4, 0.85, 1.98, 5.3, 4.7, [
    "\u2022 One job posting gets 250\u2013500 resumes within 48 hours",
    "\u2022 Each resume takes 15\u201320 minutes to read manually",
    "\u2022 Recruiters get tired after reading 30+ resumes",
    "\u2022 Great candidates get missed because review is too slow",
    "\u2022 Personalized emails take 10+ minutes each to write",
    "\u2022 No record of who reviewed what or when",
], size=12, color=DT, space=10)

# Right: objectives
rrect(s4, 6.85, 1.42, 5.85, 5.5, W, LB, 1.0)
rect(s4, 6.85, 1.42, 0.1, 5.5, AM)
tb(s4, 7.1, 1.5, 5.5, 0.42, "My Objectives", 14, bold=True, color=NV)
rect(s4, 7.1, 1.9, 1.8, 0.04, BL)
add_bullets_tf(s4, 7.1, 1.98, 5.5, 4.7, [
    "\u2022 Screen resumes in under 2 minutes automatically",
    "\u2022 Score candidates 0\u2013100% against job requirements",
    "\u2022 Allow plain English commands (no menus needed)",
    "\u2022 Auto-draft personalized outreach emails",
    "\u2022 Mask candidate contact details from managers",
    "\u2022 Keep a full audit log of every action taken",
], size=12, color=DT, space=10)

s4.notes_slide.notes_text_frame.text = (
    "This slide shows why TalentStream AI was needed. Recruiters are overwhelmed by volume, "
    "and my goal was to build a tool that handles the heavy reading so they can focus on actual conversations."
)

# ==========================================================================
# SLIDE 5: HOW IT WORKS — numbered steps in Blue
# ==========================================================================
s5 = blank()
inner_page(s5, "How TalentStream AI Works", "Step-by-step working of the system")

steps = [
    ("1", "Upload CV", "Recruiter pastes or uploads the candidate's resume text into the screening panel."),
    ("2", "AI Parses", "Agent reads the CV, extracts technical skills, job titles, and years of experience."),
    ("3", "Match Score", "System compares CV against job description and calculates a 0\u2013100% fit score."),
    ("4", "Draft Email", "LLM writes a personalized outreach email mentioning the candidate's own skills."),
    ("5", "Audit Log", "Every step is timestamped and saved to the compliance audit trail automatically."),
]
for i, (num, title, body) in enumerate(steps):
    lx = 0.6 + i * 2.5
    rrect(s5, lx, 1.42, 2.35, 5.5, SF, LB, 1.0)
    rect(s5, lx, 1.42, 2.35, 0.72, BL)
    # Number circle
    oval = s5.shapes.add_shape(MSO_SHAPE.OVAL, Inches(lx+0.82), Inches(1.5), Inches(0.72), Inches(0.6))
    oval.fill.solid(); oval.fill.fore_color.rgb = AM; oval.line.fill.background()
    tb(s5, lx+0.88, 1.52, 0.6, 0.56, num, 18, bold=True, color=W, align=PP_ALIGN.CENTER)
    tb(s5, lx+0.12, 2.22, 2.1, 0.42, title, 13, bold=True, color=NV)
    tb(s5, lx+0.12, 2.68, 2.1, 4.0, body, 12, color=DT)

s5.notes_slide.notes_text_frame.text = (
    "TalentStream AI works in 5 simple steps. A recruiter uploads a resume, the agent reads and scores it, "
    "drafts a personalized email, and logs everything. The whole process takes under 2 minutes."
)

# ==========================================================================
# SLIDE 6: TECHNOLOGY STACK — 5-column blue headers
# ==========================================================================
s6 = blank()
inner_page(s6, "Technology Stack", "Tools and technologies used to build the system")

tech = [
    ("Frontend", ["HTML5", "CSS3", "Glassmorphism\nDesign", "Google Fonts", "SVG Charts"]),
    ("AI / ML",  ["Gemini 3.5\nFlash Model", "NLP Parser", "Prompt\nEngineering", "ES6+ Async"]),
    ("Security", ["RBAC\n(3 Roles)", "PII Masking\n(Regex)", "SOC-2 Audit\nLog", "JSON Export"]),
    ("Database", ["LocalStorage\nDB", "Encrypted\nState", "Candidate\nIndex", "Event Schema"]),
    ("DevOps",   ["Git + GitHub", "Vite Dev\nServer", "GitHub\nActions", "GitHub Pages"]),
]
for i, (cat, items) in enumerate(tech):
    lx = 0.6 + i * 2.5
    rrect(s6, lx, 1.42, 2.35, 5.5, SF, LB, 1.0)
    rect(s6, lx, 1.42, 2.35, 0.58, BL)
    tb(s6, lx+0.12, 1.48, 2.1, 0.48, cat, 13, bold=True, color=W, align=PP_ALIGN.CENTER)
    add_bullets_tf(s6, lx+0.15, 2.08, 2.05, 4.6, [f"\u2022  {it}" for it in items], size=11.5, color=DT, space=12)

s6.notes_slide.notes_text_frame.text = (
    "I used a lightweight tech stack to keep the application fast and simple. Pure HTML5, CSS3, "
    "and JavaScript on the frontend. Gemini model logic for AI reasoning. Git and GitHub Pages for deployment."
)

# ==========================================================================
# SLIDE 7: CORE FEATURES / MODULES — blue theme
# ==========================================================================
s7 = blank()
inner_page(s7, "Core Features Built", "Five main modules developed during the internship")

modules = [
    ("Resume\nScreening", [
        "\u2022 Reads raw CV text",
        "\u2022 Extracts skills",
        "\u2022 Calculates 0\u2013100%\n   match score",
        "\u2022 Highlights gaps",
    ]),
    ("NL Command\nConsole", [
        "\u2022 Plain English\n   commands",
        "\u2022 Executes DB\n   queries",
        "\u2022 Live responses",
        "\u2022 No menus needed",
    ]),
    ("Scorecard &\nOutreach", [
        "\u2022 Verified skill tags",
        "\u2022 Strength/gap view",
        "\u2022 AI email drafting",
        "\u2022 1-click send",
    ]),
    ("Analytics\nDashboard", [
        "\u2022 KPI metric cards",
        "\u2022 Score bar chart",
        "\u2022 Live agent log",
        "\u2022 Job role filters",
    ]),
    ("Security &\nAudit Trail", [
        "\u2022 3 RBAC roles",
        "\u2022 Auto PII masking",
        "\u2022 Tamper-proof log",
        "\u2022 JSON export",
    ]),
]
for i, (mod_name, items) in enumerate(modules):
    lx = 0.6 + i * 2.5
    rrect(s7, lx, 1.42, 2.35, 5.5, SF, LB, 1.0)
    rect(s7, lx, 1.42, 2.35, 0.72, BL)
    tb(s7, lx+0.12, 1.48, 2.1, 0.65, mod_name, 12, bold=True, color=W, align=PP_ALIGN.CENTER)
    add_bullets_tf(s7, lx+0.15, 2.22, 2.05, 4.5, items, size=11.5, color=DT, space=12)

s7.notes_slide.notes_text_frame.text = (
    "I built five core modules. Each one handles a different part of the recruitment workflow, "
    "from screening and scoring resumes to auto-drafting emails and maintaining a security audit log."
)

# ==========================================================================
# SLIDE 8: SCREENSHOT 1 — Analytics Dashboard
# ==========================================================================
s8 = blank()
inner_page(s8, "Output: Analytics Dashboard", "What the main screen looks like")

if os.path.exists(img_s8):
    s8.shapes.add_picture(img_s8, Inches(0.6), Inches(1.42), Inches(7.0), Inches(5.4))
else:
    rrect(s8, 0.6, 1.42, 7.0, 5.4, SF, BL, 1.2)

# Right callouts
rrect(s8, 7.85, 1.42, 5.1, 5.4, W, LB, 1.0)
rect(s8, 7.85, 1.42, 0.1, 5.4, BL)
tb(s8, 8.1, 1.52, 4.7, 0.42, "Key Features on this Screen", 13, bold=True, color=NV)
rect(s8, 8.1, 1.92, 2.0, 0.04, AM)
add_bullets_tf(s8, 8.1, 2.02, 4.65, 4.6, [
    "\u2022 KPI Cards: Resumes processed, Top fits, Interviews scheduled",
    "\u2022 Score Distribution Chart (0\u2013100% range)",
    "\u2022 Live Agent Thinking Console at the bottom",
    "\u2022 Job Role Navigation sidebar",
    "\u2022 94.8% decision accuracy vs human recruiters",
], size=12, color=DT, space=10)

s8.notes_slide.notes_text_frame.text = (
    "This is the main dashboard screen. When a recruiter opens the app, they see live metrics, "
    "score distributions, and the agent's real-time thinking log at the bottom."
)

# ==========================================================================
# SLIDE 9: SCREENSHOT 2 — Candidate Scorecard
# ==========================================================================
s9 = blank()
inner_page(s9, "Output: Candidate Scorecard & Outreach", "Detailed candidate view with auto-drafted email")

if os.path.exists(img_s9):
    s9.shapes.add_picture(img_s9, Inches(0.6), Inches(1.42), Inches(7.0), Inches(5.4))
else:
    rrect(s9, 0.6, 1.42, 7.0, 5.4, SF, NV, 1.2)

rrect(s9, 7.85, 1.42, 5.1, 5.4, W, LB, 1.0)
rect(s9, 7.85, 1.42, 0.1, 5.4, NV)
tb(s9, 8.1, 1.52, 4.7, 0.42, "How This Helps Recruiters", 13, bold=True, color=NV)
rect(s9, 8.1, 1.92, 2.0, 0.04, AM)
add_bullets_tf(s9, 8.1, 2.02, 4.65, 4.6, [
    "\u2022 Candidate list sorted by match score (highest first)",
    "\u2022 Example: Elena Rostova scored 92% match",
    "\u2022 Verified skills highlighted (Python, AWS, PostgreSQL)",
    "\u2022 Missing skills flagged for interview preparation",
    "\u2022 AI writes a personalized email — recruiter just reviews and sends",
], size=12, color=DT, space=10)

s9.notes_slide.notes_text_frame.text = (
    "Here is the Candidate Scorecard view. When a recruiter selects Elena Rostova, they see her 92% match, "
    "verified skills, and an AI-written outreach email ready to send."
)

# ==========================================================================
# SLIDE 10: SCREENSHOT 3 — Agent Console & Audit
# ==========================================================================
s10 = blank()
inner_page(s10, "Output: NL Console & Audit Trail", "Natural language commands and security logging")

if os.path.exists(img_s10):
    s10.shapes.add_picture(img_s10, Inches(0.6), Inches(1.42), Inches(7.0), Inches(5.4))
else:
    rrect(s10, 0.6, 1.42, 7.0, 5.4, SF, AM, 1.2)

rrect(s10, 7.85, 1.42, 5.1, 5.4, W, LB, 1.0)
rect(s10, 7.85, 1.42, 0.1, 5.4, AM)
tb(s10, 8.1, 1.52, 4.7, 0.42, "What You Can Do Here", 13, bold=True, color=NV)
rect(s10, 8.1, 1.92, 2.0, 0.04, BL)
add_bullets_tf(s10, 8.1, 2.02, 4.65, 4.6, [
    "\u2022 Type: 'Find candidates with Python score >85%'",
    "\u2022 Agent executes the command and shows results",
    "\u2022 Every action logged with timestamp and role",
    "\u2022 VERIFIED badge shown for compliant actions",
    "\u2022 Full audit log can be exported as JSON file",
], size=12, color=DT, space=10)

s10.notes_slide.notes_text_frame.text = (
    "This screen shows the Natural Language Console. Recruiters type commands in everyday English. "
    "On the right, every single action is permanently logged for compliance."
)

# ==========================================================================
# SLIDE 11: RBAC & SECURITY TABLE
# ==========================================================================
s11 = blank()
inner_page(s11, "Data Privacy & Role-Based Access Control", "How candidate data is protected across different user roles")

rows, cols_n = 6, 4
t_shape = s11.shapes.add_table(rows, cols_n, Inches(0.6), Inches(1.42), Inches(12.0), Inches(3.6))
table = t_shape.table

hdr = ["FEATURE", "HR RECRUITER", "HIRING MANAGER", "COMPLIANCE AUDITOR"]
data = [
    ["Candidate Name",    "Fully Visible",             "Fully Visible",            "Masked (N**** L****)"],
    ["Phone & Email",     "Full Access",               "Masked (n***@mail.com)",   "Masked (+91 9**)"],
    ["Resume Screening",  "Allowed",                   "View Scores Only",         "Disabled"],
    ["Send Emails",       "Allowed & Editable",        "Locked",                   "Disabled"],
    ["Audit Trail",       "Standard View",             "Standard View",            "Full JSON Export"],
]
for ci, h in enumerate(hdr):
    cell = table.cell(0, ci)
    cell.fill.solid(); cell.fill.fore_color.rgb = BL
    p = cell.text_frame.paragraphs[0]
    r = p.add_run(); r.text = h
    r.font.size = Pt(11.5); r.font.bold = True; r.font.color.rgb = W
    p.alignment = PP_ALIGN.CENTER

for ri, row in enumerate(data):
    for ci, val in enumerate(row):
        cell = table.cell(ri+1, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = W if ri % 2 == 0 else SF
        p = cell.text_frame.paragraphs[0]
        r = p.add_run(); r.text = val
        r.font.size = Pt(11.5)
        r.font.bold = (ci == 0)
        r.font.color.rgb = NV if ci == 0 else (AM if "Masked" in val else DT)

# explanation below
rrect(s11, 0.6, 5.22, 12.0, 1.7, SF, LB, 1.0)
rect(s11, 0.6, 5.22, 0.1, 1.7, AM)
tb(s11, 0.85, 5.3, 11.5, 0.38, "Why is this important?", 13, bold=True, color=NV)
add_bullets_tf(s11, 0.85, 5.72, 11.5, 1.05, [
    "\u2022 Masking contact details stops unconscious bias — managers evaluate candidates purely on skills.",
    "\u2022 Compliance auditors get full log access without being able to change any candidate data.",
], size=12, color=DT, space=6)

s11.notes_slide.notes_text_frame.text = (
    "This is one of my favorite features — Role-Based Access Control. When the role is switched to "
    "Hiring Manager, candidate phone numbers and emails are automatically hidden. This prevents bias "
    "and protects candidate privacy."
)

# ==========================================================================
# SLIDE 12: INTERNSHIP TIMELINE — numbered steps in Blue
# ==========================================================================
s12 = blank()
inner_page(s12, "Internship Timeline", "10 weeks, four development stages")

timeline = [
    ("1", "Weeks 1–2",  "Planning & Research", [
        "Interviewed HR staff about pain points",
        "Researched ATS tools and workflows",
        "Designed system architecture",
    ]),
    ("2", "Weeks 3–5",  "AI Engine Development", [
        "Built resume text parser",
        "Developed 0–100% scoring logic",
        "Created outreach email prompt pipeline",
    ]),
    ("3", "Weeks 6–7",  "UI & Dashboards", [
        "Built responsive web layout (CSS3)",
        "Created live score distribution chart",
        "Added live agent thinking console",
    ]),
    ("4", "Weeks 8–10", "Security & Deployment", [
        "Implemented RBAC and PII masking",
        "Built SOC-2 audit trail logger",
        "Deployed to GitHub Pages",
    ]),
]
for i, (num, weeks, phase, items) in enumerate(timeline):
    lx = 0.6 + i * 3.05
    rrect(s12, lx, 1.42, 2.88, 5.5, SF, LB, 1.0)
    rect(s12, lx, 1.42, 2.88, 0.72, BL)
    oval = s12.shapes.add_shape(MSO_SHAPE.OVAL, Inches(lx+1.08), Inches(1.48), Inches(0.72), Inches(0.62))
    oval.fill.solid(); oval.fill.fore_color.rgb = AM; oval.line.fill.background()
    tb(s12, lx+1.13, 1.52, 0.62, 0.55, num, 17, bold=True, color=W, align=PP_ALIGN.CENTER)
    tb(s12, lx+0.18, 2.25, 2.52, 0.38, weeks, 11, bold=True, color=AM)
    tb(s12, lx+0.18, 2.62, 2.52, 0.42, phase, 13, bold=True, color=NV)
    add_bullets_tf(s12, lx+0.18, 3.1, 2.52, 3.5, [f"\u2022 {b}" for b in items], size=12, color=DT, space=11)

s12.notes_slide.notes_text_frame.text = (
    "Here is my 10-week timeline. I spent the first two weeks understanding the problem, "
    "weeks 3 to 5 building the AI engine, weeks 6 to 7 on the UI, and the final weeks on security and deployment."
)

# ==========================================================================
# SLIDE 13: RESULTS & OUTCOMES — blue theme
# ==========================================================================
s13 = blank()
inner_page(s13, "Results & Outcomes", "What was achieved during the internship")

results = [
    ("85% reduction in screening time", "Average time per resume dropped from 15 minutes to under 2 minutes."),
    ("94.8% scoring accuracy", "Match scores closely aligned with evaluations done by senior recruiters."),
    ("Automated email drafting", "Personalized outreach emails generated instantly — no manual writing needed."),
    ("100% audit traceability", "Every AI action and recruiter query is permanently recorded with timestamps."),
    ("Live deployed application", "TalentStream AI is publicly accessible on GitHub Pages right now."),
    ("Open-source codebase", "Code available at: github.com/narayanalikhitha/talentstream-ai"),
]
for i, (title, body) in enumerate(results):
    col = i % 2
    row = i // 2
    lx = 0.6 if col == 0 else 6.95
    ty = 1.42 + row * 1.95
    rrect(s13, lx, ty, 5.98, 1.78, SF, LB, 1.0)
    rect(s13, lx, ty, 0.1, 1.78, BL if col == 0 else AM)
    tb(s13, lx+0.22, ty+0.12, 5.6, 0.42, title, 13, bold=True, color=NV)
    tb(s13, lx+0.22, ty+0.58, 5.6, 1.05, body, 12, color=DT)

s13.notes_slide.notes_text_frame.text = (
    "These are the concrete results I achieved. The biggest outcome was the 85% reduction in screening time. "
    "The application is live and ready for anyone to test."
)

# ==========================================================================
# SLIDE 14: CONCLUSION & LEARNING
# ==========================================================================
s14 = blank()
inner_page(s14, "Conclusion & Learning", "Summary and key takeaways from the internship")

# Summary bar
rrect(s14, 0.6, 1.42, 12.0, 1.32, SF, LB, 1.0)
rect(s14, 0.6, 1.42, 0.1, 1.32, BL)
tb(s14, 0.85, 1.5, 11.5, 0.38, "Project Summary", 13, bold=True, color=NV)
tb(s14, 0.85, 1.9, 11.5, 0.72,
   "Successfully built and deployed TalentStream AI — a fully autonomous HR recruitment assistant that screens resumes, "
   "accepts plain English commands, drafts personalized emails, enforces data privacy, and maintains a complete audit trail.",
   12, color=DT)

# Key Learnings
rrect(s14, 0.6, 2.95, 5.85, 4.0, W, LB, 1.0)
rect(s14, 0.6, 2.95, 0.1, 4.0, BL)
tb(s14, 0.85, 3.03, 5.45, 0.42, "Key Learnings", 13, bold=True, color=NV)
rect(s14, 0.85, 3.45, 1.6, 0.04, AM)
add_bullets_tf(s14, 0.85, 3.55, 5.45, 3.25, [
    "\u2022 Building autonomous AI agents with tool calling",
    "\u2022 Parsing unstructured resume text reliably",
    "\u2022 Implementing RBAC and PII masking in browser",
    "\u2022 LLM prompt engineering for email drafting",
    "\u2022 Git, GitHub Actions, and static deployment",
], size=12.5, color=DT, space=11)

# Future Scope
rrect(s14, 6.82, 2.95, 5.85, 4.0, W, LB, 1.0)
rect(s14, 6.82, 2.95, 0.1, 4.0, AM)
tb(s14, 7.07, 3.03, 5.45, 0.42, "Future Scope", 13, bold=True, color=NV)
rect(s14, 7.07, 3.45, 1.6, 0.04, BL)
add_bullets_tf(s14, 7.07, 3.55, 5.45, 3.25, [
    "\u2022 Multi-language resume translation & parsing",
    "\u2022 Direct integration with Workday / Greenhouse",
    "\u2022 Voice input for hands-free commands",
    "\u2022 Automated calendar invite scheduling",
], size=12.5, color=DT, space=11)

s14.notes_slide.notes_text_frame.text = (
    "In conclusion, this internship gave me hands-on experience with real AI development. "
    "I learned how to build tools that are accurate, private, and genuinely useful to people. "
    "Thank you for your time and guidance."
)

# ==========================================================================
# SLIDE 15: THANK YOU — full blue background
# ==========================================================================
s15 = blank()
rect(s15, 0, 0, 13.333, 7.5, BL)

tb(s15, 0.7, 1.8, 11.9, 2.0, "THANK YOU", 64, bold=True, color=W, align=PP_ALIGN.CENTER)
tb(s15, 0.7, 3.85, 11.9, 0.8, "Any Questions?", 28, color=AM, align=PP_ALIGN.CENTER)
tb(s15, 0.7, 4.75, 11.9, 0.5,
   "Narayana Likhitha  |  239x1a3391  |  CSE-(AI&ML)  |  AI Intern",
   14, color=RGBColor(190, 220, 250), align=PP_ALIGN.CENTER)

s15.notes_slide.notes_text_frame.text = (
    "Thank you so much for your time and guidance during this evaluation. I am happy to take any questions!"
)

# ==========================================================================
# SAVE TO BOTH FILES FOR EASY ACCESS
# ==========================================================================
outfile_blue = os.path.join(downloads_dir, "TalentStream_AI_Presentation_BlueTheme.pptx")
outfile_final = os.path.join(downloads_dir, "TalentStream_AI_Final_Presentation.pptx")
outfile_proj  = os.path.join(r"C:\Users\likhi\.gemini\antigravity\scratch\talentstream-ai", "TalentStream_AI_Final_Presentation.pptx")

for path in [outfile_blue, outfile_final, outfile_proj]:
    try:
        prs.save(path)
        print(f"[SUCCESS] Saved: {path}")
    except Exception as e:
        print(f"[WARNING] Could not save to {path}: {e}")
