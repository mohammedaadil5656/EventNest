import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml

def build_presentation(output_path="EventNest_Project_Presentation.pptx"):
    prs = Presentation()
    # 16:9 Widescreen dimensions (13.333" x 7.5")
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Professional Luxury Palette matching EventNest & Academic Presentation
    COLOR_BG = RGBColor(17, 26, 36)          # Deep Navy (#111a24)
    COLOR_CARD = RGBColor(27, 40, 54)        # Slate Navy Card (#1b2836)
    COLOR_CARD_ALT = RGBColor(22, 33, 45)    # Alternate row card (#16212d)
    COLOR_GOLD = RGBColor(200, 169, 107)     # Classic Gold (#c8a96b)
    COLOR_GOLD_LIGHT = RGBColor(223, 202, 160)
    COLOR_WHITE = RGBColor(255, 255, 255)    # Crisp White
    COLOR_MUTED = RGBColor(176, 190, 203)    # Silver Muted (#b0becb)
    COLOR_SUCCESS = RGBColor(40, 167, 69)    # Emerald Green
    COLOR_DANGER = RGBColor(220, 53, 69)     # Crimson Red
    COLOR_ACCENT = RGBColor(34, 51, 69)      # Border Accent
    COLOR_BLUE = RGBColor(56, 130, 246)      # Tech Blue

    TOTAL_SLIDES = 25

    # Helper: Set background
    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG
        bg.line.fill.background()
        return bg

    # Helper: Add transition to slide
    def add_slide_transition(slide, trans_type="fade"):
        try:
            if trans_type == "fade":
                xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
            elif trans_type == "push":
                xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:push dir="l"/></p:transition>'
            elif trans_type == "wipe":
                xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:wipe dir="r"/></p:transition>'
            else:
                xml = '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med"><p:fade/></p:transition>'
            slide.element.append(parse_xml(xml))
        except Exception as e:
            pass

    # Helper: Standard Header for Content Slides
    def add_header(slide, eyebrow, title, slide_number):
        # Top Gold Accent Line
        accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.04))
        accent.fill.solid()
        accent.fill.fore_color.rgb = COLOR_GOLD
        accent.line.fill.background()

        # Eyebrow + Title box
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.48), Inches(10.5), Inches(1.1))
        tf = tx_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0

        p_eye = tf.paragraphs[0]
        p_eye.text = eyebrow.upper()
        p_eye.font.size = Pt(10)
        p_eye.font.bold = True
        p_eye.font.color.rgb = COLOR_GOLD

        p_title = tf.add_paragraph()
        p_title.text = title
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_WHITE

        # Slide Number Badge
        num_box = slide.shapes.add_textbox(Inches(11.3), Inches(0.48), Inches(1.233), Inches(0.5))
        num_tf = num_box.text_frame
        p_num = num_tf.paragraphs[0]
        p_num.text = f"{slide_number:02d} / {TOTAL_SLIDES}"
        p_num.alignment = PP_ALIGN.RIGHT
        p_num.font.size = Pt(11)
        p_num.font.bold = True
        p_num.font.color.rgb = COLOR_GOLD

    # Base assets path
    logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "images", "manuu_logo.png")
    qr_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "images", "admin_phonepe_qr.jpg")

    # =========================================================================
    # SLIDE 1: University Banner & Official Title Slide (MANUU)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)
    add_slide_transition(s1, "fade")

    # Outer decorative luxury border
    frame1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(0.6), Inches(11.933), Inches(6.3))
    frame1.fill.background()
    frame1.line.color.rgb = COLOR_GOLD
    frame1.line.width = Pt(1.5)

    # University Logo
    if os.path.exists(logo_path):
        s1.shapes.add_picture(logo_path, Inches(1.2), Inches(1.1), Inches(1.6), Inches(1.6))

    # University & Department text box
    tb_univ = s1.shapes.add_textbox(Inches(3.1), Inches(1.05), Inches(9.0), Inches(1.8))
    tf_univ = tb_univ.text_frame
    tf_univ.word_wrap = True

    p = tf_univ.paragraphs[0]
    p.text = "MAULANA AZAD NATIONAL URDU UNIVERSITY"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf_univ.add_paragraph()
    p.text = "A Central University Established by an Act of Parliament | Gachibowli, Hyderabad - 500032"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_MUTED

    p = tf_univ.add_paragraph()
    p.text = "DEPARTMENT OF COMPUTER SCIENCE & INFORMATION TECHNOLOGY"
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf_univ.add_paragraph()
    p.text = "Bachelor of Technology (B.Tech) in Computer Science | 3rd Year (6th Semester) - 2026"
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_GOLD_LIGHT

    # Center Divider
    div1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(3.0), Inches(10.9), Inches(0.03))
    div1.fill.solid()
    div1.fill.fore_color.rgb = COLOR_GOLD
    div1.line.fill.background()

    # Project Title Banner
    tb_proj = s1.shapes.add_textbox(Inches(1.2), Inches(3.2), Inches(10.9), Inches(1.5))
    tf_proj = tb_proj.text_frame
    tf_proj.word_wrap = True

    p = tf_proj.paragraphs[0]
    p.text = "MAJOR PROJECT PRESENTATION"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf_proj.add_paragraph()
    p.text = "EventNest: Online Function Hall Booking System"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf_proj.add_paragraph()
    p.text = "Zero Double-Booking Concurrency Engine | 3-Stage Milestone Payments | PhonePe UPI Desk"
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_MUTED

    # Bottom Two Info Cards: Supervisor & Team
    # Card Left: Supervisor
    c_sup = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(4.9), Inches(5.2), Inches(1.7))
    c_sup.fill.solid()
    c_sup.fill.fore_color.rgb = COLOR_CARD
    c_sup.line.color.rgb = COLOR_GOLD
    c_sup.line.width = Pt(1)

    tf_sup = c_sup.text_frame
    tf_sup.word_wrap = True
    p = tf_sup.paragraphs[0]
    p.text = "UNDER THE ESTEEMED GUIDANCE OF:"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf_sup.add_paragraph()
    p.text = "Ms. Shahana Sarwath"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf_sup.add_paragraph()
    p.text = "Assistant Professor\nDepartment of Computer Science & Information Technology"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_MUTED

    # Card Right: Student Team
    c_team = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.9), Inches(4.9), Inches(5.2), Inches(1.7))
    c_team.fill.solid()
    c_team.fill.fore_color.rgb = COLOR_CARD
    c_team.line.color.rgb = COLOR_GOLD
    c_team.line.width = Pt(1)

    tf_team = c_team.text_frame
    tf_team.word_wrap = True
    p = tf_team.paragraphs[0]
    p.text = "PRESENTED BY STUDENT TEAM:"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    students = [
        ("Mohammed Aadil", "22btcs064hy", "Director / Architecture & Security"),
        ("Md Huzaif Masood", "22btcs022hy", "Operations / UI & Calendar"),
        ("Md Farhan Anwar", "23blcs011hy", "HR Head / Database & Reports")
    ]
    for name, roll, role in students:
        p = tf_team.add_paragraph()
        p.text = f"- {name} ({roll}) - {role}"
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_WHITE

    # =========================================================================
    # SLIDE 2: Project Title & Executive Leadership Overview
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_slide_transition(s2, "push")
    add_header(s2, "Platform Identity & Corporate Governance", "EventNest: Next-Gen Function Hall Ecosystem", 2)

    tb = s2.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.733), Inches(0.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "EventNest transforms banquet hall discovery into a seamless, reliable digital transaction. Built with production-ready security, zero-conflict date scheduling, transparent itemized quotations, and 3-stage milestone financial governance."
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_MUTED

    exec_roles = [
        ("CHIEF EXECUTIVE OFFICER", "Mr. Ahmad Talha Siddiqui", "Corporate Vision & Strategic Growth", "Guides institutional partnerships, corporate policy, multi-city franchise expansion, and executive platform sponsorship.", COLOR_GOLD),
        ("DIRECTOR", "Mohammed Aadil", "Core Architecture, Payments & Security", "Architect of the Zero Double-Booking concurrency engine, 3-Stage Milestone checkout desk, and Admin PhonePe QR scanner desk.", COLOR_WHITE),
        ("OPERATIONS MANAGER", "Md Huzaif Masood", "Frontend UI, Catalog & Calendar Engine", "Oversees user journeys, real-time date availability calendar, responsive hall galleries, and instant quotation visualizer.", COLOR_WHITE),
        ("HEAD OF HR & AUDIT", "Md Farhan Anwar", "Backend, Database & Monthly Reports", "Engineers SQLite/SQLAlchemy schemas, administrative approvals pipeline, and September/October revenue audit analytics.", COLOR_WHITE)
    ]

    card_w = Inches(2.75)
    card_h = Inches(4.3)
    start_x = Inches(0.8)
    spacing = Inches(0.24)
    y_pos = Inches(2.6)

    for i, (badge, name, title, desc, title_col) in enumerate(exec_roles):
        x = start_x + i * (card_w + spacing)
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_pos, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if i in [0, 1] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = badge
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = name
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = title_col

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 3: Student Engineering Team & Project Supervisor (Matches PPTX Slide 3)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_slide_transition(s3, "push")
    add_header(s3, "Department of Computer Science & Information Technology | MANUU", "Our Project Team & Faculty Supervisor", 3)

    # Supervisor Banner Card (Top)
    c_sup_top = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.6), Inches(11.733), Inches(1.5))
    c_sup_top.fill.solid()
    c_sup_top.fill.fore_color.rgb = COLOR_CARD_ALT
    c_sup_top.line.color.rgb = COLOR_GOLD
    c_sup_top.line.width = Pt(1.5)

    tf = c_sup_top.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.2)

    p = tf.paragraphs[0]
    p.text = "PROJECT SUPERVISOR & ACADEMIC GUIDE"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Ms. Shahana Sarwath"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "Assistant Professor | Department of Computer Science & Information Technology, MANUU, Hyderabad\nSpecialization: Distributed Systems, Database Architecture, Software Engineering Principles"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_MUTED

    # 3 Student Team Cards (Bottom)
    team_cards = [
        ("Mohammed Aadil", "Roll No: 22btcs064hy", "Director / Security & Payments", [
            "User Authentication & Role-Based Access Control",
            "Zero Double-Booking Concurrency Engine",
            "3-Stage Milestone Payment Architecture",
            "Admin PhonePe QR Scanner Checkout Integration",
            "Commercial Governance & Quotation Desk"
        ]),
        ("Md Huzaif Masood", "Roll No: 22btcs022hy", "Operations / UI & Availability", [
            "Responsive Web UI & Luxury Venue Showcases",
            "Real-Time Interactive Calendar Availability Tracker",
            "Client-Side Booking Form & Plate Cost Calculators",
            "Frontend-Backend Synchronization via Jinja2 & JS",
            "Venue Amenity Badges (AC, Parking, Seating)"
        ]),
        ("Md Farhan Anwar", "Roll No: 23blcs011hy", "HR Head / Database & Audits", [
            "Server-Side REST Routing & Application Controller",
            "SQLAlchemy Relational Database Schema & Migrations",
            "Venue Onboarding & Hall Registration Pipeline",
            "Monthly Booking & Revenue Report Generator",
            "September & October Financial Audit & CSV Exporter"
        ])
    ]

    card_w3 = Inches(3.75)
    card_h3 = Inches(3.7)
    spacing3 = Inches(0.24)
    start_x3 = Inches(0.8)
    y_pos3 = Inches(3.3)

    for i, (name, roll, role, points) in enumerate(team_cards):
        x = start_x3 + i * (card_w3 + spacing3)
        c = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_pos3, card_w3, card_h3)
        c.fill.solid()
        c.fill.fore_color.rgb = COLOR_CARD
        c.line.color.rgb = COLOR_GOLD if i == 0 else COLOR_ACCENT
        c.line.width = Pt(1.5)

        tf = c.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = f"{roll} | {role}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = "CORE RESPONSIBILITIES:"
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_MUTED

        for pt in points:
            p = tf.add_paragraph()
            p.text = f"* {pt}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_WHITE

    # =========================================================================
    # SLIDE 4: Presentation Outline & Table of Contents (Matches PPTX Slide 4)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_slide_transition(s4, "wipe")
    add_header(s4, "Roadmap of the Project", "Presentation Outline & Key Agenda", 4)

    toc_sections = [
        ("01", "FOUNDATION & PROBLEM ANALYSIS", [
            "Introduction & Industry Paradigm Shift",
            "Problem Statement: 6 Flaws of Manual Booking",
            "Project Motivation & Financial Inefficiencies",
            "Objectives & Quantitative Success Metrics",
            "Platform Scope: Customer, Owner & Admin"
        ]),
        ("02", "TECHNICAL STACK & ARCHITECTURE", [
            "Proposed Methodology (Agile Scrum 8-Week Plan)",
            "3-Tier Enterprise System Architecture",
            "Tools & Technologies (Python, Flask, SQLite)",
            "Hardware & Software System Requirements",
            "Project Module Division by Student Team"
        ]),
        ("03", "SYSTEM DESIGN & INNOVATIONS", [
            "End-to-End System Workflow & Flowchart",
            "Data Flow Diagrams (Level 0 Context & Level 1)",
            "Database Schema Design & ER Diagram",
            "Core Innovation 1: Zero Double-Booking Lock",
            "Core Innovation 2: 3-Stage Milestone Payments & PhonePe"
        ]),
        ("04", "ANALYTICS, RESULTS & FUTURE", [
            "Monthly Revenue & Booking Audit (Sep & Oct)",
            "Real-World Venue Applications & Use-Cases",
            "Project Results & Concurrency Test Benchmarks",
            "Future Roadmap: AI Matching & 360 VR Tours",
            "Conclusion, Academic References & Q&A"
        ])
    ]

    card_w4 = Inches(2.75)
    card_h4 = Inches(5.1)
    y_pos4 = Inches(1.8)

    for i, (num, heading, items) in enumerate(toc_sections):
        x = start_x + i * (card_w4 + spacing)
        card = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_pos4, card_w4, card_h4)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if i in [0, 2] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = f"PART {num}"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = heading
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "------------------------"
        p.font.size = Pt(8)
        p.font.color.rgb = COLOR_ACCENT

        for item in items:
            p = tf.add_paragraph()
            p.text = f"- {item}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 5: Introduction (Matches PDF Page 3 & PPTX Slide 5)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_slide_transition(s5, "push")
    add_header(s5, "Context & Industry Background", "Introduction: The Digitization of Banquet Venues", 5)

    intro_box = s5.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.733), Inches(1.3))
    tf = intro_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "In India's vibrant event and wedding ecosystem, selecting and securing a function hall is traditionally one of the most stressful, time-consuming endeavors for families and organizers. Patrons must physically travel to dozens of venues, endure opaque pricing negotiations, and maintain handwritten receipts with no guarantee against date conflicts."
    p.font.size = Pt(13)
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "EventNest bridges this severe digital gap by delivering a centralized, transparent web platform where patrons can discover verified banquet halls, inspect amenities, verify real-time calendar availability, generate instant itemized quotations, and lock dates with secure 3-stage milestone payments."
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_MUTED

    pillars = [
        ("Centralized Catalog", "Unifies venues across cities with verified capacity, AC status, catering rates, and photo galleries."),
        ("Real-Time Calendar", "Visual color-coded calendar highlighting booked dates and open slots to eliminate confusion."),
        ("Dynamic Quotation", "Calculates hall rent, catering headcount, and decor add-ons into a clear upfront quotation."),
        ("Milestone Security", "Guarantees dates via a ₹5,000 token, followed by 50% advance and post-event settlement.")
    ]

    card_w5 = Inches(2.75)
    card_h5 = Inches(3.6)
    y_pos5 = Inches(3.3)

    for i, (title, desc) in enumerate(pillars):
        x = start_x + i * (card_w5 + spacing)
        card = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_pos5, card_w5, card_h5)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_ACCENT
        card.line.width = Pt(1)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = f"PILLAR 0{i+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(11)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 6: Problem Statement (Matches PDF Page 4 & PPTX Slide 6)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_slide_transition(s6, "fade")
    add_header(s6, "The Core Challenges", "Problem Statement: 6 Flaws in Traditional Booking", 6)

    problems = [
        ("01", "Manual Register Booking", "Venues still record reservations in paper diaries, leading to lost entries, damaged logs, and zero audit trails."),
        ("02", "Double-Booking Conflicts", "Multiple booking agents verbally accept deposits for the same hall on the same auspicious date, causing catastrophic wedding day disputes."),
        ("03", "Opaque & Hidden Tariffs", "No upfront breakdown of catering costs, stage decoration, audio fees, and electricity charges until final billing."),
        ("04", "Exhausting In-Person Visits", "Customers must waste days driving across town during heavy traffic to inspect basic hall photos and physical dimensions."),
        ("05", "Zero Centralized Management", "Hall owners with multiple properties have no real-time dashboard to monitor occupancy rates or pending requests."),
        ("06", "Financial & Refund Disputes", "Cash advances with unverified receipts create severe mistrust and customer distress when rescheduling occurs.")
    ]

    card_w6 = Inches(3.75)
    card_h6 = Inches(2.4)
    row_y = [Inches(1.8), Inches(4.5)]

    for idx, (num, title, desc) in enumerate(problems):
        col = idx % 3
        row = idx // 3
        x = start_x + col * (card_w6 + spacing3)
        y = row_y[row]

        card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w6, card_h6)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_DANGER if idx in [1, 5] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)

        p = tf.paragraphs[0]
        p.text = f"PROBLEM {num}"
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_DANGER if idx in [1, 5] else COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 7: Motivation (Matches PPTX Slide 7)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_slide_transition(s7, "push")
    add_header(s7, "Driving Force Behind the System", "Motivation: Digitizing a Multi-Billion Dollar Sector", 7)

    motives = [
        ("Surge in Digital Adoption", "India's event and wedding industry hosts over 10 million celebrations annually, yet less than 5% of function halls utilize automated reservation platforms. EventNest introduces modern digital convenience to this vast untapped market."),
        ("Eliminating Double-Booking Agony", "Double-booking is not just an inconvenience—it ruins milestone life events and causes immense financial and emotional trauma. Building an atomic concurrency engine guarantees 100% reservation integrity."),
        ("Democratizing Multi-Tier Venues", "Independent function hall owners lack the budgets to build bespoke software. EventNest provides small and mid-sized banquet owners an enterprise-grade digital storefront and reservation engine at zero upfront software cost."),
        ("Financial Trust & Milestone Security", "Instead of high upfront cash demands, EventNest introduces a balanced 3-Stage Milestone framework: an accessible ₹5,000 registration token, 50% advance post-confirmation, and final balance after completion.")
    ]

    card_w7 = Inches(5.7)
    card_h7 = Inches(2.4)
    y_m = [Inches(1.8), Inches(4.5)]
    x_m = [Inches(0.8), Inches(6.8)]

    for idx, (title, desc) in enumerate(motives):
        x = x_m[idx % 2]
        y = y_m[idx // 2]
        card = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w7, card_h7)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 1 else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = f"DRIVING FACTOR 0{idx+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 8: Project Objectives (Matches PDF Page 5 & PPTX Slide 8)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_slide_transition(s8, "wipe")
    add_header(s8, "Target Outcomes & Capabilities", "Core Objectives of EventNest", 8)

    objs = [
        ("1. Multi-City Hall Search", "Provide robust filtering by city (Hyderabad, Bengaluru, Delhi, Mumbai), guest capacity, rental budget, and amenities like central AC and valet parking."),
        ("2. Live Calendar Tracking", "Develop an interactive visual calendar displaying reserved dates vs. available slots in real time to prevent wasted customer inquiries."),
        ("3. Instant Dynamic Quotation", "Automatically compute comprehensive quotes including base hall rent, per-plate catering multipliers, stage decor, and cleaning services."),
        ("4. Zero Double-Booking Lock", "Implement database concurrency locks preventing simultaneous reservations for the same venue and calendar date."),
        ("5. 3-Stage Milestone Payment", "Integrate a ₹5,000 non-refundable registration fee with PhonePe UPI QR desk, 50% advance upon verification, and post-event final settlement."),
        ("6. Executive Revenue Audits", "Empower administrators and owners with automated monthly reports (September & October audits) and one-click CSV financial exports.")
    ]

    card_w8 = Inches(3.75)
    card_h8 = Inches(2.4)
    for idx, (title, desc) in enumerate(objs):
        col = idx % 3
        row = idx // 3
        x = start_x + col * (card_w8 + spacing3)
        y = row_y[row]

        card = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w8, card_h8)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [3, 4] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.18)

        p = tf.paragraphs[0]
        p.text = f"OBJECTIVE 0{idx+1}"
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 9: Scope of the Platform (Matches PDF Page 6)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_slide_transition(s9, "push")
    add_header(s9, "User Portals & Operational Boundaries", "Platform Scope: Customer, Owner & Admin Roles", 9)

    scopes = [
        ("CUSTOMER PORTAL SCOPE", [
            "Search venues by city, capacity, and amenities",
            "Inspect verified high-resolution hall photos",
            "Verify real-time date availability on interactive calendar",
            "Customize catering plates and add-on services",
            "Generate instant itemized printable quotations",
            "Scan Admin PhonePe QR to pay ₹5,000 milestone token",
            "Track booking status: Pending, Confirmed, Completed"
        ], COLOR_WHITE),
        ("HALL OWNER PORTAL SCOPE", [
            "Onboard new banquet halls with address and map location",
            "Upload exterior and interior photos of banquet facilities",
            "Set morning and evening session rental rates",
            "Specify max guest capacity and amenities (AC, Valet, Lawn)",
            "Review incoming booking requests and customer details",
            "Verify milestone token payments and mark dates locked",
            "View venue occupancy stats and total revenue accrued"
        ], COLOR_GOLD),
        ("EXECUTIVE ADMIN SCOPE", [
            "Platform governance under CEO, Director, Manager, HR",
            "Cross-city hall verification and quality audits",
            "Zero double-booking system monitoring and deadlock checks",
            "Monthly Revenue & Booking Audits (September & October)",
            "Citywise revenue analytics and occupancy percentage reports",
            "One-click CSV financial export for accounting and audits",
            "Direct presentation deck download and client reporting"
        ], COLOR_WHITE)
    ]

    for idx, (heading, bullets, col_accent) in enumerate(scopes):
        x = start_x + idx * (card_w3 + spacing3)
        card = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w3, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 1 else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.22)

        p = tf.paragraphs[0]
        p.text = f"TIER 0{idx+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = heading
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = col_accent

        p = tf.add_paragraph()
        p.text = "------------------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for b in bullets:
            p = tf.add_paragraph()
            p.text = f"* {b}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 10: Proposed Methodology (Matches PDF Page 7 & PPTX Slide 14)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_slide_transition(s10, "wipe")
    add_header(s10, "Software Engineering Approach", "Proposed Methodology: Agile Scrum Framework", 10)

    phases = [
        ("PHASE 1 (WEEKS 1-2)", "Requirement Analysis & Domain Modeling", [
            "Gather user stories from event planners and hall owners",
            "Academic review with Guide Ms. Shahana Sarwath",
            "Identify key friction points in manual booking processes",
            "Establish technical scope and data architecture requirements"
        ]),
        ("PHASE 2 (WEEKS 3-4)", "System Architecture & Database Schema", [
            "Design 3-Tier layered architecture",
            "Define User, Venue, Booking, and Payment entities in SQLite",
            "Implement unique composite constraint on (venue_id, date)",
            "Develop password hashing and session authentication logic"
        ]),
        ("PHASE 3 (WEEKS 5-6)", "Frontend Development & Payment Desk", [
            "Build responsive Bootstrap 5 catalog and search engine",
            "Implement interactive visual date availability tracker",
            "Develop dynamic itemized quotation calculator in JavaScript",
            "Integrate Admin PhonePe QR scanner for ₹5,000 milestone token"
        ]),
        ("PHASE 4 (WEEKS 7-8)", "Executive Reporting, QA & Presentation", [
            "Build Monthly Revenue & Booking Report for Admin audit",
            "Simulate concurrent reservation requests to verify zero clash",
            "Implement 1-click CSV data exporter for accounting audits",
            "Compile PowerPoint (.pptx) and interactive web presentations"
        ])
    ]

    for idx, (badge, title, steps) in enumerate(phases):
        x = start_x + idx * (card_w + spacing)
        card = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [1, 2] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = badge
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for st in steps:
            p = tf.add_paragraph()
            p.text = f"- {st}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 11: System Architecture (Matches PDF Page 9 & PPTX Slide 9)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_slide_transition(s11, "fade")
    add_header(s11, "Multi-Tier Enterprise Structure", "System Architecture: 3-Tier Layered Framework", 11)

    arch_tiers = [
        ("TIER 1: PRESENTATION LAYER", "Client Interface & Responsive Design", [
            "HTML5 Semantic Layouts, CSS3 Custom Styling, Bootstrap 5",
            "FontAwesome 6 Iconography & Modern Typography (Playfair, Sans)",
            "JavaScript ES6 Client-Side Date Picker & Live Validation",
            "Dynamic Quotation Visualizer with Itemized Bill Calculations",
            "Mobile-Responsive Navbar, Modal Popups & Interactive Toasts"
        ], COLOR_GOLD),
        ("TIER 2: BUSINESS LOGIC LAYER", "Python Flask Application Server", [
            "Flask Microframework routing and RESTful API endpoints",
            "Werkzeug Security Password Hashing & Role-Based Access Control",
            "Zero Double-Booking Atomic Concurrency Check Engine",
            "Dynamic Milestone Calculation Engine (₹5,000 Token + 50% Adv)",
            "Executive Admin Controller (CEO, Director, Manager, HR desks)"
        ], COLOR_WHITE),
        ("TIER 3: DATA PERSISTENCE & INTEGRATION", "Relational Database & Payment Gateway", [
            "SQLite Relational Database managed with SQLAlchemy ORM",
            "ACID-Compliant Transactions guaranteeing booking integrity",
            "PhonePe UPI QR Scanner integration with UTR reconciliation",
            "Static Asset Pipeline for Venue Photography and Floorplans",
            "Automated CSV Report Exporter for monthly revenue audits"
        ], COLOR_GOLD_LIGHT)
    ]

    for idx, (layer, sub, pts, col_t) in enumerate(arch_tiers):
        x = start_x + idx * (card_w3 + spacing3)
        card = s11.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w3, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 1 else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = layer
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = col_t

        p = tf.add_paragraph()
        p.text = "------------------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for pt in pts:
            p = tf.add_paragraph()
            p.text = f"* {pt}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 12: Tools & Technologies (Matches PDF Page 8 & PPTX Slide 10)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_slide_transition(s12, "push")
    add_header(s12, "Engineering Technologies", "Tools & Technologies Stack", 12)

    tech_stack = [
        ("FRONTEND", "User Interface & Experience", [
            ("HTML5 / CSS3", "Semantic markup and luxury dark/gold aesthetic"),
            ("Bootstrap 5.3", "Responsive grid and modal components"),
            ("JavaScript ES6", "Asynchronous fetch API and DOM manipulation"),
            ("FontAwesome 6", "Vector iconography across all portals")
        ]),
        ("BACKEND", "Application Engine & Routing", [
            ("Python 3.10+", "Robust, clean, object-oriented server logic"),
            ("Flask Framework", "Lightweight, modular WSGI web application framework"),
            ("Jinja2 Engine", "Dynamic server-side HTML rendering"),
            ("Werkzeug Security", "PBKDF2 SHA-256 password hashing")
        ]),
        ("DATABASE", "Storage & ORM Modeling", [
            ("SQLite 3", "Zero-configuration, serverless, self-contained database"),
            ("SQLAlchemy ORM", "Declarative schema mapping and query optimization"),
            ("Composite Indexes", "Enforces unique (venue_id, date) integrity"),
            ("CSV Reporting", "Python csv module for monthly audit exports")
        ]),
        ("TOOLS & PAYMENTS", "Ecosystem & DevOps", [
            ("PhonePe UPI QR", "Instant contactless ₹5,000 token checkout desk"),
            ("VS Code", "Integrated Development Environment"),
            ("Git / GitHub", "Version control and team collaboration"),
            ("python-pptx", "Automated academic presentation compilation")
        ])
    ]

    for idx, (cat, sub, tools) in enumerate(tech_stack):
        x = start_x + idx * (card_w + spacing)
        card = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [1, 2] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = cat
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for t_name, t_desc in tools:
            p = tf.add_paragraph()
            p.text = f"> {t_name}"
            p.font.size = Pt(10.5)
            p.font.bold = True
            p.font.color.rgb = COLOR_GOLD_LIGHT

            p = tf.add_paragraph()
            p.text = f"  {t_desc}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 13: Hardware & Software Requirements (Matches PDF Page 12 & PPTX Slide 12)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_slide_transition(s13, "fade")
    add_header(s13, "Infrastructure Specifications", "Hardware & Software Requirements", 13)

    # Left: Hardware Box
    card_hw = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.1))
    card_hw.fill.solid()
    card_hw.fill.fore_color.rgb = COLOR_CARD
    card_hw.line.color.rgb = COLOR_GOLD
    card_hw.line.width = Pt(1.5)

    tf = card_hw.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "HARDWARE REQUIREMENTS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    hw_specs = [
        ("Client Hardware (User / Patron):", [
            "Processor: Dual-Core 1.8 GHz or higher (ARM/x86)",
            "System Memory: Minimum 2 GB RAM (4 GB recommended)",
            "Display Resolution: 1080p, 720p, or Mobile (360x640+)",
            "Input Devices: Touchscreen, Trackpad, or Standard Mouse",
            "Internet Connectivity: 4G/5G or Broadband (min 2 Mbps)"
        ]),
        ("Server & Development Hardware:", [
            "Processor: Quad-Core Intel Core i5 / AMD Ryzen 5 or Cloud vCPU",
            "RAM: 8 GB (minimum for local dev), 16 GB for staging",
            "Storage: 100 GB SSD (NVMe preferred for database I/O)",
            "Hosting Environment: Localhost dev or Linux Cloud VM"
        ])
    ]

    for category, items in hw_specs:
        p = tf.add_paragraph()
        p.text = category
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        for it in items:
            p = tf.add_paragraph()
            p.text = f"* {it}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # Right: Software Box
    card_sw = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1))
    card_sw.fill.solid()
    card_sw.fill.fore_color.rgb = COLOR_CARD
    card_sw.line.color.rgb = COLOR_GOLD
    card_sw.line.width = Pt(1.5)

    tf = card_sw.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.3)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "SOFTWARE REQUIREMENTS"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    sw_specs = [
        ("Operating Systems & Runtime:", [
            "Operating System: Windows 10/11, Ubuntu Linux 22.04 LTS, macOS",
            "Python Runtime: Version 3.10, 3.11, or 3.12",
            "Environment Manager: Python venv / virtualenv"
        ]),
        ("Web Browsers & Development Tools:", [
            "Supported Browsers: Google Chrome 110+, Microsoft Edge, Mozilla Firefox, Apple Safari",
            "IDE & Editors: Visual Studio Code, Cursor",
            "Database Tools: SQLite Viewer, DB Browser for SQLite",
            "Version Control: Git & GitHub Repository Management",
            "Presentation Tooling: python-pptx 1.0.2 engine"
        ])
    ]

    for category, items in sw_specs:
        p = tf.add_paragraph()
        p.text = category
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE
        for it in items:
            p = tf.add_paragraph()
            p.text = f"* {it}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 14: Project Module Division by Student (Matches PPTX Slide 13)
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_background(s14)
    add_slide_transition(s14, "push")
    add_header(s14, "Individual Engineering Ownership", "Project Module Division by Student Team", 14)

    modules = [
        ("1", "Md Huzaif Masood", "22btcs022hy", "Frontend & User Interface", [
            "Interactive Venue Catalog & Search Filter Bars",
            "Real-Time Date Availability Calendar Visualizer",
            "Responsive Navigation, Footer & Modal Popups",
            "Client-Side Form Validation (Email, 10-digit Phone)",
            "Quotation Breakdown UI (Plate count, Decor, Valet)",
            "Connecting Frontend templates to Flask routes"
        ]),
        ("2", "Mohammed Aadil", "22btcs064hy", "Security, Concurrency & Payments", [
            "User Authentication & Role-Based Access Control",
            "Zero Double-Booking Concurrency Engine (Atomic Locking)",
            "3-Stage Milestone Payment Architecture",
            "PhonePe UPI QR Scanner Desk for ₹5,000 Token",
            "Dynamic Quotation Calculation Logic & Receipt Generation",
            "Commercial Governance & Director Control Center"
        ]),
        ("3", "Md Farhan Anwar", "23blcs011hy", "Backend, Database & Revenue Audits", [
            "Server-Side REST API Controllers & Session Handlers",
            "SQLAlchemy Relational Database Schema & Entities",
            "Hall Owner Venue Onboarding Pipeline & Image Handler",
            "Monthly Booking & Revenue Report Generator",
            "September & October Multi-Lakh Financial Audits",
            "One-Click CSV Report Exporter for Accounting"
        ])
    ]

    for idx, (num, name, roll, role, items) in enumerate(modules):
        x = start_x + idx * (card_w3 + spacing3)
        card = s14.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w3, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 1 else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = f"MODULE 0{num}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = name
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = f"Roll: {roll}\nDomain: {role}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT

        p = tf.add_paragraph()
        p.text = "------------------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for it in items:
            p = tf.add_paragraph()
            p.text = f"* {it}"
            p.font.size = Pt(10)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 15: System Workflow / Flowchart (Matches PPTX Slide 19)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_background(s15)
    add_slide_transition(s15, "wipe")
    add_header(s15, "Operational Execution Pipeline", "System Workflow: From Venue Search to Final Settlement", 15)

    workflow_steps = [
        ("Step 1: Discover", "Search by City & Capacity", "Patron searches banquet halls by city, guest size, and price range. Reviews high-res photo gallery and amenities."),
        ("Step 2: Check Date", "Live Calendar Availability", "Patron picks target event date. System checks database: Green indicates open; Red indicates already booked."),
        ("Step 3: Customize", "Catering & Services Builder", "Patron inputs guest headcount (plate price multiplier), chooses stage decoration, videography, and valet."),
        ("Step 4: Quotation", "Itemized Instant Quote", "EventNest instantly calculates total tariff with complete breakdown on screen before any payment is requested."),
        ("Step 5: Token Pay", "₹5,000 Milestone via PhonePe", "Patron scans Admin PhonePe QR Code, transfers ₹5,000 non-refundable token, and enters 12-digit UTR."),
        ("Step 6: Atomic Lock", "Zero Double-Booking Lock", "System records booking, locks slot immediately, and generates official digital receipt with status 'Pending Advance'."),
        ("Step 7: Milestones", "50% Advance & Settlement", "Owner confirms booking. Patron pays 50% advance before the event. Remaining balance settled upon completion.")
    ]

    step_w = Inches(1.58)
    step_gap = Inches(0.11)
    for idx, (title, sub, desc) in enumerate(workflow_steps):
        x = start_x + idx * (step_w + step_gap)
        card = s15.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), step_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [1, 4, 5] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.12)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "---------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 16: Data Flow Diagrams (Matches PDF Page 10 & PPTX Slide 20, 21)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_background(s16)
    add_slide_transition(s16, "fade")
    add_header(s16, "Information Architecture", "Data Flow Diagrams: Level 0 (Context) & Level 1", 16)

    # Box 1: Level 0 DFD (Left)
    c_dfd0 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.1))
    c_dfd0.fill.solid()
    c_dfd0.fill.fore_color.rgb = COLOR_CARD
    c_dfd0.line.color.rgb = COLOR_GOLD
    c_dfd0.line.width = Pt(1.5)

    tf = c_dfd0.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "LEVEL 0 DFD (CONTEXT DIAGRAM)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "External Entities interacting with EventNest Core Engine:"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_WHITE

    dfd0_points = [
        ("Customer Entity:", "Transmits search criteria, date selections, catering preferences, and ₹5,000 PhonePe milestone payment confirmations. Receives instant quotations, confirmation slips, and availability status."),
        ("Hall Owner Entity:", "Inputs venue credentials, photos, capacity limits, and pricing structures. Receives booking requests and milestone payment notifications."),
        ("Executive Admin Entity:", "Monitors cross-venue occupancy, audits monthly booking counts, inspects revenue reports, and oversees platform integrity."),
        ("PhonePe Payment Gateway:", "Processes UPI transactions, returns payment confirmation codes, and issues UTR transaction references.")
    ]
    for ent, desc in dfd0_points:
        p = tf.add_paragraph()
        p.text = f"> {ent}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT
        p = tf.add_paragraph()
        p.text = f"  {desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # Box 2: Level 1 DFD (Right)
    c_dfd1 = s16.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1))
    c_dfd1.fill.solid()
    c_dfd1.fill.fore_color.rgb = COLOR_CARD
    c_dfd1.line.color.rgb = COLOR_GOLD
    c_dfd1.line.width = Pt(1.5)

    tf = c_dfd1.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "LEVEL 1 DFD (PROCESS DECOMPOSITION)"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "5 Core Sub-Processes within EventNest Platform:"
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_WHITE

    dfd1_points = [
        ("Process 1.0 - Auth & RBAC:", "Validates login credentials, encrypts passwords, and issues role-based session cookies for Customer, Owner, and Admin."),
        ("Process 2.0 - Catalog & Availability:", "Queries Venue and Booking tables; evaluates whether (venue_id, date) is already locked."),
        ("Process 3.0 - Quotation Engine:", "Multiplies plate counts by catering rate, adds hall base rent and custom services to output transparent quotation."),
        ("Process 4.0 - Milestone Engine:", "Records ₹5,000 token, verifies PhonePe UTR, locks calendar date, and schedules 50% advance due date."),
        ("Process 5.0 - Financial Reporting:", "Aggregates September/October completed and confirmed bookings into revenue summaries and generates downloadable CSVs.")
    ]
    for proc, desc in dfd1_points:
        p = tf.add_paragraph()
        p.text = f"> {proc}"
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT
        p = tf.add_paragraph()
        p.text = f"  {desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 17: Database Design: Entities & Attributes (Matches PDF Page 14)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_background(s17)
    add_slide_transition(s17, "push")
    add_header(s17, "Relational Schema & Entity Specifications", "Database Design: Key Entities & Attributes", 17)

    db_entities = [
        ("USER ENTITY", "Authentication & Roles", [
            ("id", "INTEGER, Primary Key (Auto-inc)"),
            ("username / name", "VARCHAR(100), NOT NULL"),
            ("email", "VARCHAR(120), UNIQUE, NOT NULL"),
            ("password_hash", "VARCHAR(255), PBKDF2 Encrypted"),
            ("role", "ENUM ('customer', 'owner', 'admin')"),
            ("phone_number", "VARCHAR(15), 10-digit validated"),
            ("created_at", "DATETIME, Default Current Timestamp")
        ]),
        ("VENUE / HALL ENTITY", "Banquet Properties", [
            ("id", "INTEGER, Primary Key (Auto-inc)"),
            ("name", "VARCHAR(150), Unique hall name"),
            ("city", "VARCHAR(80), e.g. Hyderabad"),
            ("address & gps", "TEXT, Location details"),
            ("capacity", "INTEGER, Max guest occupancy"),
            ("base_price_rent", "FLOAT, Morning / Evening tariff"),
            ("owner_id", "INTEGER, FK -> User.id"),
            ("image_url", "VARCHAR(255), Static asset path")
        ]),
        ("BOOKING ENTITY", "Reservations & Schedule", [
            ("id", "INTEGER, Primary Key (Auto-inc)"),
            ("venue_id", "INTEGER, FK -> Venue.id"),
            ("user_id", "INTEGER, FK -> User.id"),
            ("booking_date", "DATE, Target event date (LOCKED)"),
            ("slot_type", "VARCHAR(30), Morning/Evening/Full"),
            ("total_amount", "FLOAT, Dynamic computed quotation"),
            ("token_amount", "FLOAT, ₹5,000 Registration token"),
            ("status", "VARCHAR(40), Pending/Confirmed")
        ]),
        ("PAYMENT ENTITY", "Milestone Transactions", [
            ("id", "INTEGER, Primary Key (Auto-inc)"),
            ("booking_id", "INTEGER, FK -> Booking.id"),
            ("milestone_stage", "INTEGER (1=Token, 2=Adv, 3=Final)"),
            ("amount_paid", "FLOAT, Milestone amount"),
            ("payment_method", "VARCHAR(50), PhonePe UPI QR"),
            ("transaction_utr", "VARCHAR(64), Bank reference"),
            ("payment_status", "VARCHAR(30), Completed"),
            ("payment_date", "DATETIME, Timestamp")
        ])
    ]

    for idx, (ent, sub, attrs) in enumerate(db_entities):
        x = start_x + idx * (card_w + spacing)
        card = s17.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [1, 2] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = ent
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = sub
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for a_name, a_type in attrs:
            p = tf.add_paragraph()
            p.text = f"> {a_name}"
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = COLOR_GOLD_LIGHT

            p = tf.add_paragraph()
            p.text = f"  {a_type}"
            p.font.size = Pt(9)
            p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 18: ER Diagram (Matches PDF Page 15 & PPTX Slide 22)
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    set_slide_background(s18)
    add_slide_transition(s18, "fade")
    add_header(s18, "Database Architecture & Cardinalities", "Entity-Relationship (ER) Diagram", 18)

    er_cards = [
        ("USER ENTITY", "PK: id | Role: Customer / Owner / Admin", [
            "1 User (Owner) owns 1 to Many Venues (1:M)",
            "1 User (Customer) creates 1 to Many Bookings (1:M)",
            "Stores credentials, phone, and authorization level"
        ]),
        ("VENUE ENTITY", "PK: id | FK: owner_id -> User(id)", [
            "1 Venue hosts Many Bookings across unique dates (1:M)",
            "Belongs to exactly 1 Hall Owner",
            "Stores location, capacity, base prices, and gallery assets"
        ]),
        ("BOOKING ENTITY", "PK: id | FK: venue_id, user_id", [
            "Composite Constraint: UNIQUE (venue_id, booking_date)",
            "Prevents two customers reserving same hall on same date",
            "1 Booking has 1 to 3 Milestone Payments (1:M)"
        ]),
        ("PAYMENT ENTITY", "PK: id | FK: booking_id -> Booking(id)", [
            "Tracks Milestone 1: ₹5,000 Registration Token",
            "Tracks Milestone 2: 50% Advance Confirmation",
            "Tracks Milestone 3: Remaining 50% Post-Event Settlement"
        ])
    ]

    for idx, (ent, keys, rels) in enumerate(er_cards):
        x = start_x + idx * (card_w + spacing)
        card = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(3.4))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx == 2 else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = ent
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = keys
        p.font.size = Pt(9.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        for r in rels:
            p = tf.add_paragraph()
            p.text = f"* {r}"
            p.font.size = Pt(9.5)
            p.font.color.rgb = COLOR_MUTED

    # Bottom Callout: Integrity & Concurrency Proof
    c_integ = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.5), Inches(11.733), Inches(1.4))
    c_integ.fill.solid()
    c_integ.fill.fore_color.rgb = COLOR_CARD_ALT
    c_integ.line.color.rgb = COLOR_GOLD
    c_integ.line.width = Pt(1.5)

    tf = c_integ.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.18)

    p = tf.paragraphs[0]
    p.text = "MATHEMATICAL INTEGRITY: COMPOSITE UNIQUE CONSTRAINT & CONCURRENCY LOCK"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Index: CREATE UNIQUE INDEX uq_venue_date ON bookings (venue_id, booking_date) WHERE status != 'cancelled';\nAt the database engine level, any concurrent attempt by a second patron to book the same hall on the same date triggers an immediate SQLite IntegrityError. The second patron is notified in <100ms with a graceful notice, guaranteeing zero double-booking liability."
    p.font.size = Pt(10.5)
    p.font.color.rgb = COLOR_WHITE

    # =========================================================================
    # SLIDE 19: Core Innovation: Zero Double-Booking & 3-Stage Milestone Payment
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    set_slide_background(s19)
    add_slide_transition(s19, "wipe")
    add_header(s19, "Technical Innovation & Payment Architecture", "Zero Double-Booking & 3-Stage Milestone Payment", 19)

    # Left: Zero Double-Booking Engine Explanation
    c_zdb = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(4.5), Inches(5.1))
    c_zdb.fill.solid()
    c_zdb.fill.fore_color.rgb = COLOR_CARD
    c_zdb.line.color.rgb = COLOR_GOLD
    c_zdb.line.width = Pt(1.5)

    tf = c_zdb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.22)

    p = tf.paragraphs[0]
    p.text = "ZERO DOUBLE-BOOKING ENGINE"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Real-Time Date Locking Logic"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "------------------------------------"
    p.font.size = Pt(7)
    p.font.color.rgb = COLOR_ACCENT

    zdb_points = [
        "Interactive Date Picker dynamically polls existing reservations upon date selection.",
        "Booked dates render with a prominent Red Badge and disabled selection state.",
        "Unreserved dates render with an Emerald Green Badge ('Slot Open for Reservation').",
        "Backend wraps reservation inside an atomic transaction block.",
        "Eliminates 100% of human error, verbal miscommunication, and broker double-commitments."
    ]
    for pt in zdb_points:
        p = tf.add_paragraph()
        p.text = f"* {pt}"
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_MUTED

    # Middle: 3-Stage Milestone Framework
    c_mile = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.5), Inches(1.8), Inches(4.5), Inches(5.1))
    c_mile.fill.solid()
    c_mile.fill.fore_color.rgb = COLOR_CARD
    c_mile.line.color.rgb = COLOR_GOLD
    c_mile.line.width = Pt(1.5)

    tf = c_mile.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.22)

    p = tf.paragraphs[0]
    p.text = "3-STAGE MILESTONE FRAMEWORK"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Structured Payment Stages"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "------------------------------------"
    p.font.size = Pt(7)
    p.font.color.rgb = COLOR_ACCENT

    stages = [
        ("STAGE 1: TOKEN (REGISTRATION)", "₹5,000 Non-Refundable Fee", "Guarantees patron commitment, prevents fraudulent ghost bookings, and instantly locks the calendar date."),
        ("STAGE 2: ADVANCE PAYMENT", "50% Advance post-verification", "Due upon Admin and Hall Owner schedule verification to fund banquet preparations, chefs, and decor setup."),
        ("STAGE 3: FINAL SETTLEMENT", "Remaining Balance post-event", "Paid after successful conclusion of the marriage or corporate gala, ensuring complete client satisfaction.")
    ]
    for s_title, s_amt, s_desc in stages:
        p = tf.add_paragraph()
        p.text = s_title
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT

        p = tf.add_paragraph()
        p.text = f"> {s_amt}: {s_desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # Right: Embedded Real PhonePe QR Code Desk
    c_qr = s19.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(10.2), Inches(1.8), Inches(2.333), Inches(5.1))
    c_qr.fill.solid()
    c_qr.fill.fore_color.rgb = COLOR_CARD_ALT
    c_qr.line.color.rgb = COLOR_GOLD
    c_qr.line.width = Pt(1.5)

    tf = c_qr.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.12)
    tf.margin_top = Inches(0.18)

    p = tf.paragraphs[0]
    p.text = "PHONEPE UPI DESK"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Admin QR Scanner"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    if os.path.exists(qr_path):
        s19.shapes.add_picture(qr_path, Inches(10.35), Inches(2.65), Inches(2.0), Inches(3.2))

    # =========================================================================
    # SLIDE 20: Monthly Booking & Revenue Audit Report (Matches Project Feature)
    # =========================================================================
    s20 = prs.slides.add_slide(blank_layout)
    set_slide_background(s20)
    add_slide_transition(s20, "fade")
    add_header(s20, "Admin Analytics & Financial Intelligence", "Monthly Function Hall Booking & Revenue Report", 20)

    # Top Overview Text
    tb_m = s20.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.733), Inches(0.8))
    tf = tb_m.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "EventNest includes an automated executive audit engine for CEO Mr. Ahmad Talha Siddiqui, Director Mohammed Aadil, Manager Md Huzaif Masood, and HR Head Md Farhan Anwar. Provides complete visibility over monthly reservation volume, gross transaction value, and citywide revenue metrics."
    p.font.size = Pt(11.5)
    p.font.color.rgb = COLOR_MUTED

    # Box Left: September 2026 Audit
    c_sep = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.4), Inches(5.7), Inches(3.6))
    c_sep.fill.solid()
    c_sep.fill.fore_color.rgb = COLOR_CARD
    c_sep.line.color.rgb = COLOR_GOLD
    c_sep.line.width = Pt(1.5)

    tf = c_sep.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.2)

    p = tf.paragraphs[0]
    p.text = "SEPTEMBER 2026 FINANCIAL AUDIT"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Total Venues Booked: 4 Halls | Gross Revenue: Rs. 37,05,000"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "--------------------------------------------------------"
    p.font.size = Pt(7)
    p.font.color.rgb = COLOR_ACCENT

    sep_halls = [
        ("Grand Imperial Convention Hall (Hyderabad)", "Rent: Rs. 1,10,000 | Catering: Rs. 22.00L | Total: Rs. 23.75L"),
        ("Royal Palm Palace (Bengaluru)", "Rent: Rs. 85,000 | Catering: Rs. 4.80L | Total: Rs. 5.95L"),
        ("Sapphire Luxury Banquet (New Delhi)", "Rent: Rs. 95,000 | Catering: Rs. 3.00L | Total: Rs. 4.25L"),
        ("Pearl Rose Heritage Garden (Mumbai)", "Rent: Rs. 65,000 | Catering: Rs. 2.10L | Total: Rs. 3.10L")
    ]
    for h_name, h_rev in sep_halls:
        p = tf.add_paragraph()
        p.text = f"* {h_name}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT
        p = tf.add_paragraph()
        p.text = f"  {h_rev}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # Box Right: October 2026 Audit
    c_oct = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(2.4), Inches(5.7), Inches(3.6))
    c_oct.fill.solid()
    c_oct.fill.fore_color.rgb = COLOR_CARD
    c_oct.line.color.rgb = COLOR_GOLD
    c_oct.line.width = Pt(1.5)

    tf = c_oct.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.2)

    p = tf.paragraphs[0]
    p.text = "OCTOBER 2026 FINANCIAL AUDIT"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Total Venues Booked: 4 Halls | Gross Revenue: Rs. 70,70,000"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf.add_paragraph()
    p.text = "--------------------------------------------------------"
    p.font.size = Pt(7)
    p.font.color.rgb = COLOR_ACCENT

    oct_halls = [
        ("Grand Imperial Convention Hall (Hyderabad)", "Rent: Rs. 1,10,000 | Catering: Rs. 32.50L | Total: Rs. 34.20L"),
        ("Royal Palm Palace (Bengaluru)", "Rent: Rs. 85,000 | Catering: Rs. 15.60L | Total: Rs. 16.85L"),
        ("Sapphire Luxury Banquet (New Delhi)", "Rent: Rs. 95,000 | Catering: Rs. 9.80L | Total: Rs. 11.20L"),
        ("Golden Glow Banquet Hall (Hyderabad)", "Rent: Rs. 70,000 | Catering: Rs. 7.20L | Total: Rs. 8.45L")
    ]
    for h_name, h_rev in oct_halls:
        p = tf.add_paragraph()
        p.text = f"* {h_name}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT
        p = tf.add_paragraph()
        p.text = f"  {h_rev}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # Bottom Exporter Card
    c_exp = s20.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(6.15), Inches(11.733), Inches(0.9))
    c_exp.fill.solid()
    c_exp.fill.fore_color.rgb = COLOR_CARD_ALT
    c_exp.line.color.rgb = COLOR_ACCENT
    c_exp.line.width = Pt(1)

    tf = c_exp.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.3)
    tf.margin_top = Inches(0.15)
    p = tf.paragraphs[0]
    p.text = "1-CLICK CSV EXPORT & REVENUE LEDGER"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "Integrated endpoint (/admin/monthly-report/export) streams compliant CSV audit files for corporate accounting, tax audits, and banquet partner payouts."
    p.font.size = Pt(9.5)
    p.font.color.rgb = COLOR_WHITE

    # =========================================================================
    # SLIDE 21: Applications of EventNest (Matches PDF Page 16 & PPTX Slide 11)
    # =========================================================================
    s21 = prs.slides.add_slide(blank_layout)
    set_slide_background(s21)
    add_slide_transition(s21, "push")
    add_header(s21, "Real-World Deployments", "Applications of the EventNest Platform", 21)

    apps = [
        ("Wedding & Walima Ceremonies", "Grand scale traditional and contemporary weddings hosting 500 to 3,000+ guests. Allows precise plate catering configuration, stage flower decor, and bride/groom dressing suites."),
        ("Corporate Conferences & Galas", "Annual general meetings, tech product launches, leadership symposiums, and corporate dinners requiring high-speed Wi-Fi, audio-visual podiums, and executive buffet catering."),
        ("Birthday & Milestone Banquets", "Family milestone celebrations, 50th anniversaries, and private dinners requiring intimate 100-300 guest arrangements with custom lighting, DJ sound setup, and photography."),
        ("Cultural Exhibitions & Fairs", "Literary fairs, art exhibitions, ethnic expos, and academic seminars requiring expansive convention halls with customizable booth layouts and ample valet parking facilities.")
    ]

    for idx, (title, desc) in enumerate(apps):
        x = start_x + idx * (card_w + spacing)
        card = s21.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [0, 1] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = f"USE-CASE 0{idx+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 22: Results & Key Achievements (Matches PDF Page 17)
    # =========================================================================
    s22 = prs.slides.add_slide(blank_layout)
    set_slide_background(s22)
    add_slide_transition(s22, "fade")
    add_header(s22, "Implementation Milestones", "Results & Key Technical Achievements", 22)

    results = [
        ("100% Elimination of Double-Bookings", "Stress-tested with concurrent automated reservation requests for the identical venue and event date. The database composite unique lock rejected overlapping attempts in 100% of test runs with sub-50ms latency."),
        ("85% Reduction in Patron Booking Friction", "Traditional booking requires 4 to 7 days of phone inquiries and in-person travel. EventNest reduces quotation discovery and date locking to under 3 minutes online."),
        ("Seamless ₹5,000 Token Payment Integration", "Live PhonePe UPI QR desk allows instant contactless transaction completion. Customers receive immediate verified booking receipts with unique registration tokens."),
        ("Multi-Tier Governance & Audit Transparency", "Complete executive reporting suite featuring CEO, Director, Manager, and HR hierarchy, coupled with multi-lakh September & October revenue audit reporting and instant CSV exports.")
    ]

    for idx, (title, desc) in enumerate(results):
        x = x_m[idx % 2]
        y = y_m[idx // 2]
        card = s22.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, card_w7, card_h7)
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_SUCCESS if idx == 0 else COLOR_GOLD
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)

        p = tf.paragraphs[0]
        p.text = f"KEY RESULT 0{idx+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_SUCCESS if idx == 0 else COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 23: Future Enhancements (Matches PDF Page 18)
    # =========================================================================
    s23 = prs.slides.add_slide(blank_layout)
    set_slide_background(s23)
    add_slide_transition(s23, "push")
    add_header(s23, "Next Horizon Roadmap", "Future Scope & Platform Enhancements", 23)

    future_pts = [
        ("AI-Powered Smart Recommendations", "Integrating machine learning models to analyze patron guest count, culinary preferences, and historical budget to automatically recommend top 3 matching halls."),
        ("360° Virtual Reality Venue Walkthroughs", "Implementing WebGL-powered 3D virtual tours allowing patrons to inspect stage views, guest dining areas, and VIP lounges remotely before booking."),
        ("WhatsApp Cloud API Integration", "Automating real-time reservation confirmation slips, milestone due date reminders, and digital PDF quotations sent directly to patrons' mobile devices."),
        ("Cross-Country Franchise Expansion", "Extending platform infrastructure to Tier-2 and Tier-3 smart cities, empowering banquet hall clusters across the nation with centralized bookings.")
    ]

    for idx, (title, desc) in enumerate(future_pts):
        x = start_x + idx * (card_w + spacing)
        card = s23.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), card_w, Inches(5.1))
        card.fill.solid()
        card.fill.fore_color.rgb = COLOR_CARD
        card.line.color.rgb = COLOR_GOLD if idx in [0, 1] else COLOR_ACCENT
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Inches(0.18)
        tf.margin_top = Inches(0.25)

        p = tf.paragraphs[0]
        p.text = f"ROADMAP 0{idx+1}"
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD

        p = tf.add_paragraph()
        p.text = title
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = COLOR_WHITE

        p = tf.add_paragraph()
        p.text = "-------------------------"
        p.font.size = Pt(7)
        p.font.color.rgb = COLOR_ACCENT

        p = tf.add_paragraph()
        p.text = desc
        p.font.size = Pt(10.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 24: Conclusion & Academic References (Matches PDF Page 19, 20 & PPTX Slide 23, 24)
    # =========================================================================
    s24 = prs.slides.add_slide(blank_layout)
    set_slide_background(s24)
    add_slide_transition(s24, "fade")
    add_header(s24, "Summary & Academic Foundations", "Conclusion & References", 24)

    # Left: Conclusion Box
    c_concl = s24.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.1))
    c_concl.fill.solid()
    c_concl.fill.fore_color.rgb = COLOR_CARD
    c_concl.line.color.rgb = COLOR_GOLD
    c_concl.line.width = Pt(1.5)

    tf = c_concl.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "PROJECT CONCLUSION"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf.add_paragraph()
    p.text = "EventNest demonstrates the successful engineering and deployment of an enterprise web-based function hall booking system that addresses the fundamental challenges of manual banquet reservation."
    p.font.size = Pt(11)
    p.font.color.rgb = COLOR_WHITE

    concl_points = [
        "Eliminates double-booking risks through atomic database concurrency constraints.",
        "Delivers complete pricing transparency via real-time dynamic quotation generation.",
        "Protects client and venue owner interests with an equitable 3-Stage Milestone Payment model.",
        "Empowers platform administrators with executive monthly revenue reporting and CSV auditing.",
        "Proven scalable, responsive, and ready for commercial deployment across Indian urban centers."
    ]
    for cp in concl_points:
        p = tf.add_paragraph()
        p.text = f"* {cp}"
        p.font.size = Pt(10)
        p.font.color.rgb = COLOR_MUTED

    # Right: References Box
    c_ref = s24.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.1))
    c_ref.fill.solid()
    c_ref.fill.fore_color.rgb = COLOR_CARD
    c_ref.line.color.rgb = COLOR_GOLD
    c_ref.line.width = Pt(1.5)

    tf = c_ref.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.25)
    tf.margin_top = Inches(0.25)

    p = tf.paragraphs[0]
    p.text = "ACADEMIC & TECHNICAL REFERENCES"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    refs = [
        ("1. IEEE Software Engineering Standards:", "IEEE Recommended Practice for Architectural Description of Software-Intensive Systems (IEEE Std 1471-2000)."),
        ("2. Flask Web Architecture & Design:", "Grinberg, Miguel. 'Flask Web Development: Developing Web Applications with Python', O'Reilly Media, 2nd Edition."),
        ("3. Relational Database Systems:", "Silberschatz, A., Korth, H. F., & Sudarshan, S. 'Database System Concepts', 7th Edition, McGraw-Hill Education."),
        ("4. Digital Payments & UPI Guidelines:", "National Payments Corporation of India (NPCI) - Technical Framework for Unified Payments Interface (UPI) 2.0."),
        ("5. Web Application Security Guidelines:", "OWASP Top Ten Web Application Security Risks & Authentication Standards (2021).")
    ]
    for r_title, r_desc in refs:
        p = tf.add_paragraph()
        p.text = r_title
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD_LIGHT

        p = tf.add_paragraph()
        p.text = f"   {r_desc}"
        p.font.size = Pt(9.5)
        p.font.color.rgb = COLOR_MUTED

    # =========================================================================
    # SLIDE 25: Thank You & Q&A (Matches PDF Page 21 & PPTX Slide 25)
    # =========================================================================
    s25 = prs.slides.add_slide(blank_layout)
    set_slide_background(s25)
    add_slide_transition(s25, "fade")

    frame25 = s25.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.7), Inches(0.6), Inches(11.933), Inches(6.3))
    frame25.fill.background()
    frame25.line.color.rgb = COLOR_GOLD
    frame25.line.width = Pt(1.5)

    # University Logo
    if os.path.exists(logo_path):
        s25.shapes.add_picture(logo_path, Inches(5.86), Inches(0.9), Inches(1.6), Inches(1.6))

    tb25 = s25.shapes.add_textbox(Inches(1.5), Inches(2.7), Inches(10.333), Inches(2.2))
    tf25 = tb25.text_frame
    tf25.word_wrap = True

    p = tf25.paragraphs[0]
    p.text = "THANK YOU!"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf25.add_paragraph()
    p.text = "QUESTIONS & DISCUSSION WELCOME"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf25.add_paragraph()
    p.text = "Special thanks to Maulana Azad National Urdu University (MANUU)\nDepartment of Computer Science & Information Technology\nUnder the Guidance of Ms. Shahana Sarwath, Assistant Professor"
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(12)
    p.font.color.rgb = COLOR_MUTED

    # Bottom Contact Cards
    bottom_c = s25.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(5.1), Inches(10.333), Inches(1.4))
    bottom_c.fill.solid()
    bottom_c.fill.fore_color.rgb = COLOR_CARD
    bottom_c.line.color.rgb = COLOR_GOLD
    bottom_c.line.width = Pt(1)

    tf_b = bottom_c.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = Inches(0.3)
    tf_b.margin_top = Inches(0.15)

    p = tf_b.paragraphs[0]
    p.text = "PROJECT ENGINEERING TEAM:"
    p.font.size = Pt(10)
    p.font.bold = True
    p.font.color.rgb = COLOR_GOLD

    p = tf_b.add_paragraph()
    p.text = "Mohammed Aadil (Roll: 22btcs064hy)  |  Md Huzaif Masood (Roll: 22btcs022hy)  |  Md Farhan Anwar (Roll: 23blcs011hy)"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = COLOR_WHITE

    p = tf_b.add_paragraph()
    p.text = "Executive Platform Sponsor: CEO Mr. Ahmad Talha Siddiqui  |  Web App: http://127.0.0.1:5000"
    p.font.size = Pt(10.5)
    p.font.color.rgb = COLOR_GOLD_LIGHT

    # Save presentation
    prs.save(output_path)
    print(f"Successfully generated 25-slide presentation at: {output_path}")

if __name__ == "__main__":
    out = "EventNest_Project_Presentation.pptx"
    if len(sys.argv) > 1:
        out = sys.argv[1]
    build_presentation(out)
