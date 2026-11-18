import html
import re
from html import escape as xml_escape
from io import BytesIO
from urllib.parse import quote_plus

import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from pypdf import PdfReader

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import Image as ReportLabImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


st.set_page_config(
    page_title="RoleFit | Resume Tools",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)


SKILLS = [
    "Python", "Java", "JavaScript", "TypeScript", "C++", "C#",
    "SQL", "Git", "Django", "Flask", "FastAPI", "React", "HTML",
    "CSS", "AWS", "Azure", "Docker", "Linux", "REST API",
    "unit testing", "data analysis", "machine learning",
    "communication", "leadership", "project management", "Excel",
    "customer service", "customer support", "sales", "writing",
    "teamwork", "problem solving", "time management", "marketing",
    "research", "budgeting", "organization", "presentation",
    "negotiation", "attention to detail", "conflict resolution",
    "client relations", "event planning", "inventory management",
    "bookkeeping", "training", "administrative support", "scheduling",
    "data entry", "record keeping", "invoicing", "calendar management",
    "Microsoft Office", "Google Workspace", "social media",
    "content creation", "SEO", "graphic design", "copywriting",
    "email marketing", "logistics", "quality assurance", "hospitality",
    "food safety",
]


SKILL_ALIASES = {
    "customer service": ["customer care", "client support"],
    "communication": ["verbal communication", "written communication"],
    "project management": ["project coordination", "managing projects"],
    "data entry": ["data input"],
    "record keeping": ["record management", "maintaining records"],
    "scheduling": ["appointment scheduling"],
    "teamwork": ["collaboration", "working in a team"],
    "problem solving": ["troubleshooting", "issue resolution"],
    "attention to detail": ["detail oriented", "strong attention to detail"],
    "administrative support": ["office administration", "admin support"],
}


SKILL_SUGGESTIONS = {
    "Python": "Build a small script or app, then add it to your portfolio.",
    "SQL": "Practice queries using a sample database and save a few examples.",
    "Git": "Put a small project in a Git repository and practice commits and branches.",
    "Excel": "Practice formulas, sorting, filters, and pivot tables on sample data.",
    "communication": "Practice writing a clear work update or presenting a short topic.",
    "customer service": "Practice responding to common customer questions and documenting resolutions.",
    "project management": "Plan a small project with tasks, deadlines, and progress updates.",
    "data analysis": "Analyze a small dataset and summarize your findings in a chart or report.",
    "writing": "Create a short writing sample related to the kind of role you want.",
    "social media": "Draft a week of sample posts for a business or cause.",
    "presentation": "Prepare a short presentation and practice delivering it clearly.",
    "marketing": "Create a sample campaign plan for a product, service, or cause.",
    "graphic design": "Create a few sample designs and collect them in a small portfolio.",
    "sales": "Practice a product pitch and handling common customer questions.",
    "leadership": "Coordinate a small task or group project and document your contribution.",
}


SKILL_RESOURCES = {
    "Python": ("Official Python Tutorial", "https://docs.python.org/3/tutorial/"),
    "SQL": ("PostgreSQL SQL Tutorial", "https://www.postgresql.org/docs/current/tutorial.html"),
    "Git": ("Pro Git Book", "https://git-scm.com/book/en/v2"),
    "machine learning": (
        "Google Machine Learning Crash Course",
        "https://developers.google.com/machine-learning/crash-course",
    ),
    "Excel": (
        "Microsoft Excel Support and Training",
        "https://support.microsoft.com/en-us/excel",
    ),
    "writing": (
        "Purdue OWL Writing Resources",
        "https://owl.purdue.edu/owl/general_writing/index.html",
    ),
    "communication": (
        "Purdue OWL Writing Process Resources",
        "https://owl.purdue.edu/owl/general_writing/the_writing_process/index.html",
    ),
}


TEMPLATES = {
    "Professional": {
        "accent": "243B64",
        "description": "Centered name and contact details with a traditional section order.",
        "order": [
            ("Professional Summary", "summary"),
            ("Skills", "skills"),
            ("Work Experience", "experience"),
            ("Education", "education"),
            ("Projects", "projects"),
        ],
    },
    "Modern": {
        "accent": "5146C8",
        "description": "A colored header banner and a contemporary section order.",
        "order": [
            ("Professional Summary", "summary"),
            ("Work Experience", "experience"),
            ("Projects", "projects"),
            ("Skills", "skills"),
            ("Education", "education"),
        ],
    },
    "Skills-first": {
        "accent": "28745D",
        "description": "Places skills near the top, followed by summary and projects.",
        "order": [
            ("Skills", "skills"),
            ("Professional Summary", "summary"),
            ("Projects", "projects"),
            ("Work Experience", "experience"),
            ("Education", "education"),
        ],
    },
    "Photo profile": {
        "accent": "315C81",
        "description": "Places your contact details beside an optional profile photo.",
        "order": [
            ("Professional Summary", "summary"),
            ("Skills", "skills"),
            ("Work Experience", "experience"),
            ("Education", "education"),
            ("Projects", "projects"),
        ],
    },
}


def extract_pdf_text(pdf_file):
    """Extract text from every page of an uploaded PDF."""
    reader = PdfReader(pdf_file)
    return "\n".join(page.extract_text() or "" for page in reader.pages)


def normalize_text(text):
    """Lowercase text and replace punctuation with spaces."""
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def get_skill_phrases(skill):
    """Return a skill and its alternate phrases."""
    return [skill, *SKILL_ALIASES.get(skill, [])]


def find_skills(text):
    """Find known skills and alternate phrases in text."""
    normalized_text = f" {normalize_text(text)} "
    found_skills = set()

    for skill in SKILLS:
        for phrase in get_skill_phrases(skill):
            if f" {normalize_text(phrase)} " in normalized_text:
                found_skills.add(skill)
                break

    return found_skills


def find_skill_evidence(text, skills):
    """Find a resume sentence containing each matched skill."""
    sentences = re.split(r"(?<=[.!?])\s+|\n+", text)
    evidence = {}

    for skill in skills:
        for sentence in sentences:
            normalized_sentence = f" {normalize_text(sentence)} "

            if any(
                f" {normalize_text(phrase)} " in normalized_sentence
                for phrase in get_skill_phrases(skill)
            ):
                evidence[skill] = sentence.strip()
                break

    return evidence


def check_resume_format(text):
    """Check for common resume sections and contact details."""
    normalized_text = normalize_text(text)

    section_names = {
        "Experience": ["experience", "work experience", "employment history"],
        "Education": ["education", "academic background"],
        "Skills": ["skills", "technical skills", "core competencies"],
    }

    checks = {
        section: any(normalize_text(name) in normalized_text for name in names)
        for section, names in section_names.items()
    }

    checks["Email address"] = bool(
        re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", text)
    )
    checks["Phone number"] = bool(
        re.search(r"(?<!\d)(?:\+?\d[\d\s().-]{7,}\d)(?!\d)", text)
    )

    return checks


def get_study_link(skill):
    """Return a curated study link or a search link for a skill."""
    if skill in SKILL_RESOURCES:
        return SKILL_RESOURCES[skill]

    query = quote_plus(f"{skill} beginner tutorial learning")
    return (
        f"Search study materials for {skill}",
        f"https://www.google.com/search?q={query}",
    )


def set_cell_background(cell, color):
    """Set a Word table cell background color."""
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), color)
    properties.append(shading)


def add_header_to_cell(cell, resume_data, accent, white_text=False):
    """Add name and contact details inside a Word table cell."""
    name_paragraph = cell.paragraphs[0]
    name_run = name_paragraph.add_run(resume_data["name"] or "Your Name")
    name_run.bold = True
    name_run.font.name = "Arial"
    name_run.font.size = Pt(17)
    name_run.font.color.rgb = RGBColor.from_string(
        "FFFFFF" if white_text else accent
    )

    contact_items = [
        resume_data["email"],
        resume_data["phone"],
        resume_data["location"],
        resume_data["linkedin"],
        resume_data["portfolio"],
    ]
    contact_line = " | ".join(
        item.strip() for item in contact_items if item.strip()
    )

    if contact_line:
        contact_paragraph = cell.add_paragraph()
        contact_run = contact_paragraph.add_run(contact_line)
        contact_run.font.size = Pt(8)
        contact_run.font.color.rgb = RGBColor.from_string(
            "F2F4FF" if white_text else "596579"
        )


def add_plain_header(document, resume_data, accent, centered=False):
    """Add a regular name-and-contact header to a Word document."""
    name_paragraph = document.add_paragraph()

    if centered:
        name_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    name_run = name_paragraph.add_run(resume_data["name"] or "Your Name")
    name_run.bold = True
    name_run.font.name = "Georgia" if centered else "Arial"
    name_run.font.size = Pt(18)
    name_run.font.color.rgb = RGBColor.from_string(accent)

    contact_items = [
        resume_data["email"],
        resume_data["phone"],
        resume_data["location"],
        resume_data["linkedin"],
        resume_data["portfolio"],
    ]
    contact_line = " | ".join(
        item.strip() for item in contact_items if item.strip()
    )

    if contact_line:
        contact = document.add_paragraph(contact_line)
        if centered:
            contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact.paragraph_format.space_after = Pt(4)


def add_resume_header(document, resume_data, template_name, photo_bytes):
    """Add a distinct header for the chosen resume layout."""
    accent = TEMPLATES[template_name]["accent"]

    if template_name == "Modern":
        table = document.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        set_cell_background(cell, accent)
        add_header_to_cell(cell, resume_data, accent, white_text=True)
        document.add_paragraph()
        return

    if template_name == "Photo profile":
        table = document.add_table(rows=1, cols=2)
        table.columns[0].width = Inches(5.4)
        table.columns[1].width = Inches(1.4)

        left_cell = table.cell(0, 0)
        right_cell = table.cell(0, 1)
        add_header_to_cell(left_cell, resume_data, accent)

        if photo_bytes:
            image_paragraph = right_cell.paragraphs[0]
            image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            image_paragraph.add_run().add_picture(
                BytesIO(photo_bytes),
                width=Inches(1.0),
            )

        document.add_paragraph()
        return

    add_plain_header(
        document,
        resume_data,
        accent,
        centered=(template_name == "Professional"),
    )
    document.add_paragraph()


def add_docx_section(document, title, text, accent, template_name):
    """Add a compact styled heading and text to a Word document."""
    if not text.strip():
        return

    heading_text = title.upper() if template_name == "Modern" else title
    heading = document.add_heading(heading_text, level=1)
    heading.paragraph_format.space_before = Pt(7)
    heading.paragraph_format.space_after = Pt(3)
    heading.paragraph_format.keep_with_next = True

    for run in heading.runs:
        run.font.color.rgb = RGBColor.from_string(accent)
        run.font.size = Pt(10)
        run.font.bold = True

    for line in text.splitlines():
        line = line.strip()
        if line:
            paragraph = document.add_paragraph(line)
            paragraph.paragraph_format.space_after = Pt(1)
            paragraph.paragraph_format.line_spacing = 1.0


def create_resume_docx(resume_data, template_name, photo_bytes=None):
    """Create a compact editable Word resume."""
    template = TEMPLATES[template_name]
    document = Document()

    section = document.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.45)
    section.left_margin = Inches(0.55)
    section.right_margin = Inches(0.55)

    normal_style = document.styles["Normal"]
    normal_style.font.name = "Arial"
    normal_style.font.size = Pt(9)
    normal_style.paragraph_format.space_after = Pt(2)
    normal_style.paragraph_format.line_spacing = 1.0

    add_resume_header(document, resume_data, template_name, photo_bytes)

    for title, key in template["order"]:
        add_docx_section(
            document,
            title,
            resume_data[key],
            template["accent"],
            template_name,
        )

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output.getvalue()


def create_resume_pdf(resume_data, template_name, photo_bytes=None):
    """Create a one-page PDF resume, or raise an error if it spills to page two."""
    template = TEMPLATES[template_name]
    accent_hex = f"#{template['accent']}"
    accent_color = colors.HexColor(accent_hex)
    buffer = BytesIO()

    pdf = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=0.55 * inch,
        leftMargin=0.55 * inch,
        topMargin=0.4 * inch,
        bottomMargin=0.4 * inch,
        title=f"{resume_data['name'] or 'Resume'}",
    )

    name_style = ParagraphStyle(
        "ResumeName",
        fontName="Helvetica-Bold",
        fontSize=17,
        leading=19,
        textColor=accent_color,
        alignment=TA_LEFT,
        spaceAfter=3,
    )
    contact_style = ParagraphStyle(
        "ResumeContact",
        fontName="Helvetica",
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#596579"),
    )
    white_name_style = ParagraphStyle(
        "WhiteResumeName",
        parent=name_style,
        textColor=colors.white,
    )
    white_contact_style = ParagraphStyle(
        "WhiteResumeContact",
        parent=contact_style,
        textColor=colors.HexColor("#F2F4FF"),
    )
    heading_style = ParagraphStyle(
        "ResumeHeading",
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=11,
        textColor=accent_color,
        spaceBefore=6,
        spaceAfter=2,
        keepWithNext=True,
    )
    body_style = ParagraphStyle(
        "ResumeBody",
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
        textColor=colors.HexColor("#202939"),
        spaceAfter=2,
    )

    contact_items = [
        resume_data["email"],
        resume_data["phone"],
        resume_data["location"],
        resume_data["linkedin"],
        resume_data["portfolio"],
    ]
    contact_line = " | ".join(
        item.strip() for item in contact_items if item.strip()
    )

    safe_name = xml_escape(resume_data["name"] or "Your Name")
    safe_contact = xml_escape(contact_line)
    story = []

    if template_name == "Modern":
        header_content = [
            Paragraph(safe_name, white_name_style),
            Paragraph(safe_contact, white_contact_style) if safe_contact else "",
        ]
        header = Table([[header_content]], colWidths=[7.1 * inch])
        header.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), accent_color),
                    ("LEFTPADDING", (0, 0), (-1, -1), 12),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                    ("TOPPADDING", (0, 0), (-1, -1), 10),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
                ]
            )
        )
        story.extend([header, Spacer(1, 7)])

    elif template_name == "Photo profile":
        if photo_bytes:
            photo = ReportLabImage(
                BytesIO(photo_bytes),
                width=0.95 * inch,
                height=1.1 * inch,
            )
            header_data = [[
                [
                    Paragraph(safe_name, name_style),
                    Paragraph(safe_contact, contact_style) if safe_contact else "",
                ],
                photo,
            ]]
            header = Table(
                header_data,
                colWidths=[5.9 * inch, 1.1 * inch],
            )
            header.setStyle(
                TableStyle(
                    [
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 4),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ]
                )
            )
            story.extend([header, Spacer(1, 6)])
        else:
            story.append(Paragraph(safe_name, name_style))
            if safe_contact:
                story.append(Paragraph(safe_contact, contact_style))
            story.append(Spacer(1, 6))

    else:
        centered_style = ParagraphStyle(
            "CenteredResumeName",
            parent=name_style,
            fontName="Times-Bold" if template_name == "Professional" else "Helvetica-Bold",
            fontSize=18,
            alignment=TA_CENTER if template_name == "Professional" else TA_LEFT,
        )
        contact_alignment = (
            TA_CENTER if template_name == "Professional" else TA_LEFT
        )
        aligned_contact_style = ParagraphStyle(
            "AlignedContact",
            parent=contact_style,
            alignment=contact_alignment,
        )
        story.append(Paragraph(safe_name, centered_style))
        if safe_contact:
            story.append(Paragraph(safe_contact, aligned_contact_style))
        story.append(Spacer(1, 5))

    for title, key in template["order"]:
        content = resume_data[key].strip()
        if not content:
            continue

        heading_text = title.upper() if template_name == "Modern" else title
        story.append(Paragraph(xml_escape(heading_text), heading_style))

        for line in content.splitlines():
            line = line.strip()
            if line:
                story.append(
                    Paragraph(xml_escape(line), body_style)
                )

    pdf.build(story)
    pdf_bytes = buffer.getvalue()

    page_count = len(PdfReader(BytesIO(pdf_bytes)).pages)
    if page_count != 1:
        raise ValueError(
            "This content does not fit on one page. Shorten some sections and try again."
        )

    return pdf_bytes


def show_analyzer():
    st.markdown(
        """
        <div class="hero">
            <div class="eyebrow">RESUME ANALYZER</div>
            <h1>See how your resume fits the role</h1>
            <p>Compare your experience with a job description and find practical next steps.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Analyzer settings")
        resume_source = st.radio("Resume input", ["Paste text", "Upload PDF"])
        threshold = st.slider(
            "Target match score",
            min_value=0,
            max_value=100,
            value=60,
            format="%d%%",
        )

    with st.form("resume_match_form"):
        resume_column, job_column = st.columns(2, gap="large")
        resume_text = ""
        pdf_file = None

        with resume_column:
            st.subheader("Your resume")

            if resume_source == "Paste text":
                resume_text = st.text_area(
                    "Paste resume text",
                    height=260,
                    placeholder="Paste the text from your resume here...",
                    label_visibility="collapsed",
                )
            else:
                pdf_file = st.file_uploader(
                    "Upload a resume PDF",
                    type=["pdf"],
                )

        with job_column:
            st.subheader("Job description")
            job_description = st.text_area(
                "Paste job description",
                height=260,
                placeholder="Paste the job posting here...",
                label_visibility="collapsed",
            )

        submitted = st.form_submit_button(
            "Analyze my fit",
            type="primary",
            use_container_width=True,
        )

    if not submitted:
        st.info("Add a resume and job description, then select **Analyze my fit**.")
        return

    if resume_source == "Upload PDF" and pdf_file is not None:
        try:
            resume_text = extract_pdf_text(pdf_file)
        except Exception as error:
            st.error(f"Could not read the PDF: {error}")
            return

    if not resume_text.strip():
        st.error("Add your resume text or upload a text-based PDF.")
        return

    if not job_description.strip():
        st.error("Paste a job description.")
        return

    resume_skills = find_skills(resume_text)
    job_skills = find_skills(job_description)
    matched_skills = resume_skills & job_skills
    missing_skills = job_skills - resume_skills

    score = (
        round(len(matched_skills) / len(job_skills) * 100)
        if job_skills
        else 0
    )

    evidence = find_skill_evidence(resume_text, matched_skills)
    resume_checks = check_resume_format(resume_text)

    st.divider()
    st.header("Analysis")

    score_column, matched_column, gaps_column = st.columns(3)

    with score_column:
        st.metric("Estimated skill match", f"{score}%")
        st.progress(score)

    with matched_column:
        st.metric("Skills matched", len(matched_skills))

    with gaps_column:
        st.metric("Skills to review", len(missing_skills))

    if not job_skills:
        st.info("No skills from the app's skill list were found in the job description.")
    elif score >= threshold:
        st.success(f"You reached your target of {threshold}%.")
    else:
        st.warning(f"You are below your target of {threshold}%.")

    overview_tab, gaps_tab, checklist_tab, report_tab = st.tabs(
        ["Overview", "Skill gaps & learning", "Resume checks", "Download report"]
    )

    with overview_tab:
        st.subheader("Matched skills")
        if matched_skills:
            for skill in sorted(matched_skills):
                sentence = evidence.get(skill, "No supporting sentence found.")
                with st.container(border=True):
                    st.markdown(f"**{skill}**")
                    st.caption(f"Resume evidence: “{sentence}”")
        else:
            st.info("No matching skills were found.")

        st.caption(
            "This is a simple skill overlap estimate, not a real ATS score "
            "or hiring prediction."
        )

    with gaps_tab:
        st.subheader("Skills to review")
        if missing_skills:
            for skill in sorted(missing_skills):
                suggestion = SKILL_SUGGESTIONS.get(
                    skill,
                    "Find a small practice task, complete it, and save evidence of what you learned.",
                )
                resource_title, resource_url = get_study_link(skill)

                with st.container(border=True):
                    st.markdown(f"**{skill}**")
                    st.write(suggestion)
                    st.markdown(f"[{resource_title}]({resource_url})")
        else:
            st.success("No missing skills from the app's skill list were detected.")

    with checklist_tab:
        st.subheader("Basic resume checks")
        st.caption(
            "These checks look for common text. They do not guarantee how an ATS will read or rank a resume."
        )

        checklist_columns = st.columns(2)
        for index, (item, found) in enumerate(resume_checks.items()):
            with checklist_columns[index % 2]:
                with st.container(border=True):
                    st.write(
                        f"{'✅' if found else '⚠️'} "
                        f"{item} {'detected' if found else 'not detected'}"
                    )

    with report_tab:
        report_lines = [
            "ROLEFIT RESUME MATCH REPORT",
            "",
            f"Estimated skill match: {score}%",
            f"Target score: {threshold}%",
            "",
            "MATCHED SKILLS AND RESUME EVIDENCE",
        ]

        if matched_skills:
            for skill in sorted(matched_skills):
                report_lines.append(
                    f"- {skill}: {evidence.get(skill, 'No supporting sentence found.')}"
                )
        else:
            report_lines.append("- No matching skills found.")

        report_lines.extend(["", "SKILLS TO REVIEW"])
        report_lines.extend(
            [f"- {skill}" for skill in sorted(missing_skills)]
            or ["- No missing skills found."]
        )
        report_lines.extend(["", "BASIC RESUME CHECKS"])

        for item, found in resume_checks.items():
            report_lines.append(
                f"- {item}: {'Detected' if found else 'Not detected'}"
            )

        st.download_button(
            "Download analysis report",
            data="\n".join(report_lines),
            file_name="rolefit_resume_report.txt",
            mime="text/plain",
        )


def show_resume_builder():
    st.markdown(
        """
        <div class="hero">
            <div class="eyebrow">RESUME BUILDER</div>
            <h1>Choose a layout and build your resume</h1>
            <p>Your selected layout is shown here and updates the preview as you edit.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    template_name = st.radio(
        "Choose a resume layout",
        list(TEMPLATES.keys()),
        horizontal=True,
        key="builder_template",
    )

    st.success(
        f"Selected layout: {template_name} — "
        f"{TEMPLATES[template_name]['description']}"
    )

    include_photo = False
    if template_name == "Photo profile":
        include_photo = st.checkbox(
            "Include a profile photo",
            key="builder_include_photo",
        )

    form_column, preview_column = st.columns([1.05, 0.95], gap="large")

    with form_column:
        st.subheader("Contact information")

        name = st.text_input("Name", key="builder_name")
        email = st.text_input("Email", key="builder_email")
        phone = st.text_input("Phone", key="builder_phone")
        location = st.text_input("Location", key="builder_location")
        linkedin = st.text_input("LinkedIn URL", key="builder_linkedin")
        portfolio = st.text_input("Portfolio or website URL", key="builder_portfolio")

        photo_bytes = None
        if template_name == "Photo profile" and include_photo:
            photo_file = st.file_uploader(
                "Upload a photo",
                type=["png", "jpg", "jpeg"],
                key="builder_photo_file",
            )
            if photo_file is not None:
                photo_bytes = photo_file.getvalue()

            st.caption(
                "A photo is optional. Some employers and application systems prefer resumes without photos."
            )

        st.subheader("Resume content")

        summary = st.text_area(
            "Professional summary",
            placeholder="Briefly describe your experience and strengths.",
            height=90,
            key="builder_summary",
        )
        skills = st.text_area(
            "Skills",
            placeholder="Enter skills separated by commas or on separate lines.",
            height=75,
            key="builder_skills",
        )
        experience = st.text_area(
            "Work experience",
            placeholder=(
                "Enter each role and dates. Put each achievement or "
                "responsibility on a separate line."
            ),
            height=145,
            key="builder_experience",
        )
        education = st.text_area(
            "Education",
            placeholder="Enter qualification, institution, and dates.",
            height=75,
            key="builder_education",
        )
        projects = st.text_area(
            "Projects",
            placeholder="Enter project names and what you did on each one.",
            height=90,
            key="builder_projects",
        )

    resume_data = {
        "name": name,
        "email": email,
        "phone": phone,
        "location": location,
        "linkedin": linkedin,
        "portfolio": portfolio,
        "summary": summary,
        "skills": skills,
        "experience": experience,
        "education": education,
        "projects": projects,
    }

    all_resume_text = "\n".join(resume_data.values())
    character_count = len(all_resume_text)
    line_count = sum(
        len(value.splitlines())
        for value in resume_data.values()
        if value.strip()
    )

    character_limit = 1800 if template_name == "Photo profile" else 2200
    line_limit = 28 if template_name == "Photo profile" else 32
    fits_one_page_budget = (
        character_count <= character_limit
        and line_count <= line_limit
    )

    with preview_column:
        st.subheader(f"Live preview · {template_name}")
        accent = f"#{TEMPLATES[template_name]['accent']}"
        safe_name = html.escape(name or "Your Name")
        contact_items = [email, phone, location, linkedin, portfolio]
        safe_contact = " · ".join(
            html.escape(item.strip())
            for item in contact_items
            if item.strip()
        )

        with st.container(border=True):
            if template_name == "Modern":
                st.markdown(
                    f"""
                    <div class="modern-preview" style="background:{accent};">
                        <div class="preview-kicker">CURRICULUM VITAE</div>
                        <div class="preview-name">{safe_name}</div>
                        <div class="preview-contact">{safe_contact}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            elif template_name == "Professional":
                st.markdown(
                    f"""
                    <div class="professional-preview">
                        <div class="preview-name">{safe_name}</div>
                        <div class="preview-contact">{safe_contact}</div>
                        <div class="preview-rule"></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            elif template_name == "Skills-first":
                st.markdown(
                    f"""
                    <div class="skills-preview" style="border-left-color:{accent};">
                        <div class="preview-kicker">CAREER PROFILE</div>
                        <div class="preview-name">{safe_name}</div>
                        <div class="preview-contact">{safe_contact}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                photo_column, info_column = st.columns([0.32, 0.68])
                with photo_column:
                    if photo_bytes:
                        st.image(photo_bytes, use_container_width=True)
                    else:
                        st.markdown(
                            '<div class="photo-placeholder">Optional<br>photo</div>',
                            unsafe_allow_html=True,
                        )
                with info_column:
                    st.markdown(
                        f"""
                        <div class="preview-kicker">PROFILE</div>
                        <div class="preview-name">{safe_name}</div>
                        <div class="preview-contact">{safe_contact}</div>
                        """,
                        unsafe_allow_html=True,
                    )

            for section_title, key in TEMPLATES[template_name]["order"]:
                content = resume_data[key]
                if content.strip():
                    st.markdown(
                        f'<div class="preview-section-title" '
                        f'style="color:{accent};">{html.escape(section_title)}</div>',
                        unsafe_allow_html=True,
                    )
                    st.text(content)

        st.caption(
            f"One-page guide: {character_count}/{character_limit} characters "
            f"and {line_count}/{line_limit} lines."
        )

        if not fits_one_page_budget:
            st.warning(
                "Shorten the content to fit the one-page limit before downloading."
            )

        has_content = any(value.strip() for value in resume_data.values())

        if has_content and fits_one_page_budget:
            try:
                docx_bytes = create_resume_docx(
                    resume_data,
                    template_name,
                    photo_bytes,
                )
                st.download_button(
                    "Download Word resume",
                    data=docx_bytes,
                    file_name=f"{template_name.lower().replace(' ', '_')}_resume.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    type="primary",
                    use_container_width=True,
                )
            except Exception as error:
                st.error(f"Could not create the Word document: {error}")

            try:
                pdf_bytes = create_resume_pdf(
                    resume_data,
                    template_name,
                    photo_bytes,
                )
                st.download_button(
                    "Download one-page PDF",
                    data=pdf_bytes,
                    file_name=f"{template_name.lower().replace(' ', '_')}_resume.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as error:
                st.warning(str(error))

        elif not has_content:
            st.info("Enter some details to see your resume preview.")


st.markdown(
    """
    <style>
        :root {
            --ink: #18243a;
            --muted: #66748a;
            --line: #e4e9f2;
            --brand: #5b55e7;
            --canvas: #f4f6fb;
        }

        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(ellipse at 15% 0%, #eef0ff 0, transparent 34%),
                var(--canvas);
        }

        .block-container {
            max-width: 1260px;
            padding-top: 1.5rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background: #172238;
            border-right: 1px solid #26334a;
        }

        [data-testid="stSidebar"] * {
            color: #eef2ff;
        }

        [data-testid="stSidebar"] hr {
            border-color: #34415a;
        }

        [data-testid="stSidebar"] div.stButton > button {
            border: 1px solid #394760;
            background: #202d45;
            color: #eef2ff;
            text-align: left;
        }

        [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            border-color: #817bff;
            background: #5b55e7;
            color: #ffffff;
        }

        .brand-mark, .eyebrow, .preview-kicker {
            color: #716aff;
            font-size: 0.73rem;
            font-weight: 800;
            letter-spacing: 0.15em;
        }

        [data-testid="stSidebar"] .brand-mark {
            color: #aaa6ff;
        }

        .hero {
            padding: 2rem 2.2rem;
            margin: 0.5rem 0 1.5rem 0;
            border: 1px solid #e1e5f5;
            border-radius: 24px;
            background: linear-gradient(120deg, #ffffff 0%, #f1f0ff 56%, #eaf8f4 100%);
            box-shadow: 0 14px 40px rgba(36, 47, 84, 0.06);
        }

        .hero h1 {
            max-width: 760px;
            margin: 0.45rem 0 0.55rem 0;
            color: var(--ink);
            font-size: clamp(2rem, 4vw, 2.8rem);
            font-weight: 750;
            letter-spacing: -0.055em;
            line-height: 1.12;
        }

        .hero p {
            max-width: 720px;
            margin: 0;
            color: var(--muted);
            font-size: 1.04rem;
            line-height: 1.6;
        }

        div[data-testid="stMetric"],
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border: 1px solid var(--line);
            border-radius: 16px;
            background: #ffffff;
            box-shadow: 0 8px 24px rgba(36, 47, 84, 0.04);
        }

        div[data-testid="stMetric"] {
            padding: 1.1rem 1.2rem;
        }

        div[data-testid="stTextArea"] textarea,
        div[data-testid="stTextInput"] input {
            border: 1px solid #dce2ec;
            border-radius: 12px;
            background: #ffffff;
        }

        div[data-testid="stFileUploader"] {
            padding: 0.8rem;
            border: 1px dashed #c7cede;
            border-radius: 14px;
            background: #fbfcff;
        }

        div.stButton > button,
        div.stDownloadButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            min-height: 2.9rem;
            border-radius: 12px;
            font-weight: 700;
        }

        div[data-testid="stRadio"] > div {
            gap: 0.5rem;
        }

        .modern-preview {
            padding: 1.25rem;
            margin-bottom: 1rem;
            border-radius: 12px;
            color: white;
        }

        .modern-preview .preview-kicker,
        .modern-preview .preview-contact,
        .modern-preview .preview-name {
            color: white;
        }

        .professional-preview {
            padding: 0.8rem 0.4rem 0.2rem 0.4rem;
            text-align: center;
            font-family: Georgia, serif;
        }

        .skills-preview {
            padding: 0.5rem 0.8rem 0.8rem 1rem;
            border-left: 5px solid #28745D;
        }

        .preview-name {
            color: #1b2942;
            font-size: 1.6rem;
            font-weight: 750;
            line-height: 1.2;
            overflow-wrap: anywhere;
        }

        .professional-preview .preview-name {
            font-family: Georgia, serif;
            font-size: 1.8rem;
        }

        .preview-contact {
            margin-top: 0.35rem;
            color: #69758a;
            font-size: 0.76rem;
            overflow-wrap: anywhere;
        }

        .preview-rule {
            height: 2px;
            margin: 0.8rem auto 0 auto;
            background: #243B64;
        }

        .photo-placeholder {
            display: flex;
            min-height: 110px;
            align-items: center;
            justify-content: center;
            border: 1px dashed #aab5c7;
            border-radius: 12px;
            background: #f4f6fb;
            color: #66748a;
            text-align: center;
            font-size: 0.8rem;
        }

        .preview-section-title {
            margin-top: 1rem;
            margin-bottom: 0.25rem;
            font-size: 0.9rem;
            font-weight: 750;
        }

        @media (max-width: 700px) {
            .hero {
                padding: 1.4rem;
                border-radius: 18px;
            }

            .hero h1 {
                font-size: 1.9rem;
            }

            .block-container {
                padding-top: 1rem;
            }
        }
    </style>
    """,
    unsafe_allow_html=True,
)


with st.sidebar:
    st.markdown('<div class="brand-mark">ROLEFIT</div>', unsafe_allow_html=True)
    st.title("Resume tools")
    st.caption("Analyze a role match or build a resume.")

    if "current_page" not in st.session_state:
        st.session_state.current_page = "analyzer"

    st.divider()

    if st.button(
        "Resume Analyzer",
        type="primary" if st.session_state.current_page == "analyzer" else "secondary",
        use_container_width=True,
    ):
        st.session_state.current_page = "analyzer"
        st.rerun()

    if st.button(
        "Build a resume",
        type="primary" if st.session_state.current_page == "builder" else "secondary",
        use_container_width=True,
    ):
        st.session_state.current_page = "builder"
        st.rerun()


if st.session_state.current_page == "builder":
    show_resume_builder()
else:
    show_analyzer()