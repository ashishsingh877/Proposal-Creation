"""
AI-Powered Proposal Generator
Personalises a DPDPA consulting proposal template (.pptx) for any new company.

Slide 4  – AI rewrites company description + scope (word-limited to prevent overflow)
Slide 5  – Fully rebuilt with professional layout using questionnaire data
Slide 11 – AI rewrites operating model paragraph
Slide 12, 14, 19 – Company name auto-replaced
Slide 17 – AI rewrites all 6 Data Lifecycle sections
"""

import io, json, re
import streamlit as st
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from docx import Document as DocxDocument
from groq import Groq

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="AI Proposal Generator", page_icon="📊", layout="wide")
st.title("📊 AI Proposal Generator")
st.markdown(
    "Upload the **template PPTX** + **Pre-Scoping Questionnaire (.docx)**, enter company "
    "details and your Groq key — the app personalises every slide while keeping design intact."
)

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key    = st.text_input("Groq API Key", type="password", placeholder="gsk_...")
    st.markdown("---")
    st.header("🏢 Company Details")
    company_name    = st.text_input("Full Company Name",           placeholder="SGD Pharma India Pvt. Ltd.")
    company_short   = st.text_input("Abbreviation / Short Name",   placeholder="SGD")
    company_website = st.text_input("Website URL",                 placeholder="https://www.sgd-pharma.com")
    st.markdown("---")
    st.header("📁 Files")
    uploaded_ppt    = st.file_uploader("Template PPTX",                    type=["pptx"])
    uploaded_docx   = st.file_uploader("Pre-Scoping Questionnaire (.docx)", type=["docx"])
    st.markdown("---")
    generate_btn = st.button("🚀 Generate Proposal", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────
# CONSTANTS – design tokens matching template palette
# ─────────────────────────────────────────────────────────────
FONT        = "Aptos"
C_NAVY      = RGBColor(0x1F, 0x38, 0x64)   # dark navy
C_NAVY2     = RGBColor(0x2E, 0x5E, 0x9A)   # mid-blue
C_TEAL      = RGBColor(0x1A, 0x65, 0x70)   # teal
C_ORANGE    = RGBColor(0xE8, 0x83, 0x3A)   # template accent
C_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
C_LIGHT_BG  = RGBColor(0xEB, 0xF2, 0xF7)   # left panel bg
C_DIVIDER   = RGBColor(0xCC, 0xD9, 0xE8)
C_BODY_DARK = RGBColor(0x22, 0x22, 0x33)
C_BODY_LITE = RGBColor(0xCB, 0xDC, 0xF0)   # light text on dark bg
TAG_COLORS  = [
    RGBColor(0x2E, 0x5E, 0x9A),
    RGBColor(0x1A, 0x65, 0x70),
    RGBColor(0x3A, 0x7D, 0xB4),
    RGBColor(0x17, 0x55, 0x65),
    RGBColor(0x45, 0x82, 0xBB),
]

# ─────────────────────────────────────────────────────────────
# PPTX DRAW UTILITIES
# ─────────────────────────────────────────────────────────────
def _set_tf(tf, text, font_size, bold=False, italic=False,
            color=None, align=PP_ALIGN.LEFT, word_wrap=True,
            pad=Inches(0.07)):
    tf.word_wrap = word_wrap
    tf.margin_left   = pad
    tf.margin_right  = pad
    tf.margin_top    = Inches(0.04)
    tf.margin_bottom = Inches(0.04)
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(font_size)
    run.font.bold   = bold
    run.font.italic = italic
    if color:
        run.font.color.rgb = color
    return tf

def add_rect(slide, x, y, w, h, fill, line=None, line_w=0.5, radius=False):
    """Add a filled rectangle (or rounded rect)."""
    shape_type = 5 if radius else 1   # 5=rounded, 1=rectangle
    shp = slide.shapes.add_shape(shape_type, x, y, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    return shp

def add_label(slide, x, y, w, h, text, fill, text_color=C_WHITE,
              font_size=10, bold=False, align=PP_ALIGN.LEFT,
              radius=False, pad=Inches(0.08)):
    """Filled box with single-line text."""
    shp = add_rect(slide, x, y, w, h, fill, radius=radius)
    _set_tf(shp.text_frame, text, font_size, bold=bold,
            color=text_color, align=align, pad=pad)
    return shp

def add_textbox(slide, x, y, w, h, text, font_size=9, bold=False,
                italic=False, color=C_BODY_DARK, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(x, y, w, h)
    _set_tf(txb.text_frame, text, font_size, bold=bold, italic=italic,
            color=color, align=align)
    return txb

def add_multiline_box(slide, x, y, w, h, lines, font_size=8.5,
                      color=C_WHITE, bullet_char="▪  "):
    """Text box with multiple bulleted lines."""
    txb = slide.shapes.add_textbox(x, y, w, h)
    tf = txb.text_frame
    tf.word_wrap = True
    tf.margin_left   = Inches(0.05)
    tf.margin_right  = Inches(0.05)
    tf.margin_top    = Inches(0.03)
    tf.margin_bottom = Inches(0.03)
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.LEFT
        run = para.add_run()
        run.text = f"{bullet_char}{line}"
        run.font.name = FONT
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
    return txb

def hr(slide, x, y, w, color=C_DIVIDER):
    """Thin horizontal rule."""
    shp = slide.shapes.add_shape(1, x, y, w, Pt(1))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()
    return shp

# ─────────────────────────────────────────────────────────────
# SLIDE 5 FULL REBUILD
# ─────────────────────────────────────────────────────────────
def rebuild_slide5(slide, company_name: str, company_short: str, info: dict):
    """Remove all existing shapes and draw a fresh professional Slide 5."""

    # ── 1. Wipe all shapes ───────────────────────────────────
    for shp in list(slide.shapes):
        shp._element.getparent().remove(shp._element)

    # ── 2. Dimensions ────────────────────────────────────────
    SW = Inches(13.33)
    SH = Inches(7.5)
    MARGIN = Inches(0.2)

    LP_X  = MARGIN
    LP_Y  = Inches(1.15)
    LP_W  = Inches(4.85)
    LP_H  = SH - LP_Y - MARGIN
    RP_X  = LP_X + LP_W + MARGIN
    RP_Y  = LP_Y
    RP_W  = SW - RP_X - MARGIN
    RP_H  = LP_H

    # ── 3. Slide title ───────────────────────────────────────
    # Teal title strip
    add_rect(slide, Inches(0), Inches(0), SW, Inches(1.1), C_NAVY)
    add_textbox(slide, Inches(0.3), Inches(0.18), Inches(10), Inches(0.75),
                "Scope of Review (High-Level)",
                font_size=28, bold=True, color=C_WHITE)

    # ── 4. LEFT PANEL background ─────────────────────────────
    add_rect(slide, LP_X, LP_Y, LP_W, LP_H, C_LIGHT_BG,
             line=C_DIVIDER, line_w=0.5)

    # ─ 4a. Org Overview section ──────────────────────────────
    ORG_Y = LP_Y + Inches(0.12)
    add_textbox(slide, LP_X + Inches(0.15), ORG_Y, LP_W - Inches(0.3), Inches(0.32),
                "Organizational Overview",
                font_size=12, bold=True, color=C_NAVY)

    # Employee count — full-width banner (handles short numbers AND long descriptions)
    emp       = info.get("employee_count", "—").strip()
    short_emp = len(emp) <= 8          # e.g. "1,200+"
    emp_font  = 36 if short_emp else 14

    # Dark banner spanning full panel width
    BAN_Y = ORG_Y + Inches(0.38)
    BAN_H = Inches(0.68) if short_emp else Inches(0.56)
    add_rect(slide, LP_X, BAN_Y, LP_W, BAN_H, C_NAVY)
    add_textbox(slide, LP_X + Inches(0.15), BAN_Y + Inches(0.04),
                LP_W - Inches(0.3), BAN_H - Inches(0.08),
                emp, font_size=emp_font, bold=True,
                color=C_WHITE, align=PP_ALIGN.CENTER)

    # Sub-label: "Employee Strength"
    SUB_Y = BAN_Y + BAN_H + Inches(0.04)
    add_textbox(slide, LP_X + Inches(0.15), SUB_Y,
                LP_W - Inches(0.3), Inches(0.22),
                "Employee Strength",
                font_size=8, color=RGBColor(0x55, 0x66, 0x77),
                align=PP_ALIGN.CENTER)

    # Italic review note
    note_y = SUB_Y + Inches(0.27)
    add_textbox(slide,
                LP_X + Inches(0.15), note_y,
                LP_W - Inches(0.3), Inches(0.65),
                "Review coverage will be limited to the identified in-scope "
                "functions of the parent organization.",
                font_size=8.5, italic=True, color=RGBColor(0x44,0x55,0x66))

    # ─ 4b. Personal Data Stored section ──────────────────────
    PST_Y = note_y + Inches(0.75)
    hr(slide, LP_X + Inches(0.1), PST_Y, LP_W - Inches(0.2))
    PST_Y += Inches(0.06)

    add_label(slide,
              LP_X, PST_Y, LP_W, Inches(0.36),
              C_NAVY2, bold=True, font_size=10, align=PP_ALIGN.LEFT,
              text="  🗄  Personal Data Stored And Hosted", pad=Inches(0.08))

    hosting_txt = _build_hosting_text(info)
    add_textbox(slide,
                LP_X + Inches(0.12), PST_Y + Inches(0.4),
                LP_W - Inches(0.24), Inches(0.75),
                hosting_txt, font_size=8.5, color=C_BODY_DARK)

    # ─ 4c. Departments section ───────────────────────────────
    DEPT_Y = PST_Y + Inches(1.22)
    hr(slide, LP_X + Inches(0.1), DEPT_Y, LP_W - Inches(0.2))
    DEPT_Y += Inches(0.06)

    add_label(slide,
              LP_X, DEPT_Y, LP_W, Inches(0.36),
              C_TEAL, bold=True, font_size=10, align=PP_ALIGN.LEFT,
              text="  🏢  Departments & Sub-Processes", pad=Inches(0.08))

    depts = info.get("departments", [])
    n_dept = len(depts)
    # Join with ", " for inline display; wrap naturally inside the text box
    dept_text = (f"Audit coverage includes assessment of data handling practices "
                 f"within {n_dept} departments such as: {', '.join(depts)}.")
    add_textbox(slide,
                LP_X + Inches(0.12), DEPT_Y + Inches(0.4),
                LP_W - Inches(0.24), Inches(1.3),
                dept_text, font_size=8.5, color=C_BODY_DARK)

    # ── 5. RIGHT PANEL background ────────────────────────────
    add_rect(slide, RP_X, RP_Y, RP_W, RP_H, C_NAVY)

    # Header bar
    HDR_H = Inches(0.42)
    add_label(slide,
              RP_X, RP_Y, RP_W, HDR_H,
              RGBColor(0x0E, 0x20, 0x40),
              bold=True, font_size=11, align=PP_ALIGN.CENTER,
              text="Business Operations & Digital Infrastructure")

    # ─ 5a. Business line tags ─────────────────────────────────
    blines = [b for b in info.get("core_business_lines", []) if b][:5]
    if blines:
        TAG_H   = Inches(0.38)
        TAG_Y   = RP_Y + HDR_H + Inches(0.1)
        gap     = Inches(0.06)
        tag_w   = (RP_W - gap * (len(blines) + 1)) / len(blines)
        for i, line in enumerate(blines):
            tx = RP_X + gap + i * (tag_w + gap)
            add_label(slide,
                      tx, TAG_Y, tag_w, TAG_H,
                      TAG_COLORS[i % len(TAG_COLORS)],
                      bold=True, font_size=8, align=PP_ALIGN.CENTER,
                      text=line[:28], radius=True)

    # Audit note
    NOTE_Y = RP_Y + HDR_H + Inches(0.58)
    add_textbox(slide, RP_X + Inches(0.15), NOTE_Y, RP_W - Inches(0.3), Inches(0.34),
                "Audit coverage includes review of collection, usage, storage, transfer, "
                "retention and deletion practices across business landscape.",
                font_size=7.5, color=C_BODY_LITE, align=PP_ALIGN.CENTER)

    # ─ 5b. Data types + subjects grid ────────────────────────
    GRID_Y  = NOTE_Y + Inches(0.38)
    GRID_H  = Inches(2.55)
    col_w   = (RP_W - Inches(0.45)) / 2

    # --- Critical Data Types column ---
    DT_X = RP_X + Inches(0.15)
    add_label(slide,
              DT_X, GRID_Y, col_w, Inches(0.34),
              C_NAVY2, bold=True, font_size=9, align=PP_ALIGN.CENTER,
              text="Critical Data Types")

    dtypes = [d for d in info.get("data_types", []) if d][:6]
    card_cols  = 2
    card_w     = (col_w - Inches(0.08)) / card_cols
    card_h     = Inches(0.54)
    card_gap   = Inches(0.05)
    card_start = GRID_Y + Inches(0.38)
    for i, dt in enumerate(dtypes):
        row = i // card_cols
        col = i % card_cols
        cx  = DT_X + col * (card_w + card_gap)
        cy  = card_start + row * (card_h + card_gap)
        add_label(slide,
                  cx, cy, card_w, card_h,
                  RGBColor(0x15, 0x28, 0x48),
                  text=dt[:22], font_size=7.5, align=PP_ALIGN.CENTER,
                  line=RGBColor(0x3A, 0x6A, 0xB0), line_w=0.5)

    # --- Data Subjects column ---
    DS_X = DT_X + col_w + Inches(0.15)
    add_label(slide,
              DS_X, GRID_Y, col_w, Inches(0.34),
              C_TEAL, bold=True, font_size=9, align=PP_ALIGN.CENTER,
              text="Categories of Data Subjects")

    subjects = [s for s in info.get("data_subjects", []) if s][:6]
    for i, subj in enumerate(subjects):
        row = i // card_cols
        col = i % card_cols
        cx  = DS_X + col * (card_w + card_gap)
        cy  = card_start + row * (card_h + card_gap)
        add_label(slide,
                  cx, cy, card_w, card_h,
                  RGBColor(0x10, 0x28, 0x38),
                  text=subj[:22], font_size=7.5, align=PP_ALIGN.CENTER,
                  line=RGBColor(0x1F, 0x6B, 0x75), line_w=0.5)

    # ─ 5c. Application Ecosystem ─────────────────────────────
    APP_Y = GRID_Y + GRID_H + Inches(0.08)
    add_label(slide,
              RP_X + Inches(0.15), APP_Y, RP_W - Inches(0.3), Inches(0.33),
              C_NAVY2, bold=True, font_size=9.5, align=PP_ALIGN.LEFT,
              text="  03   Application Ecosystem", pad=Inches(0.08))

    apps = [a for a in info.get("applications", []) if a][:7]
    apps_str = ", ".join(apps) if apps else "ERP, CRM, HRMS"
    add_textbox(slide,
                RP_X + Inches(0.15), APP_Y + Inches(0.36),
                RP_W - Inches(0.3), Inches(0.52),
                f"{company_short} utilizes core enterprise applications including {apps_str}.",
                font_size=8.5, color=C_BODY_LITE)

    # ─ 5d. Customer Facing Interfaces ────────────────────────
    CF_Y = APP_Y + Inches(0.93)
    add_label(slide,
              RP_X + Inches(0.15), CF_Y, RP_W - Inches(0.3), Inches(0.33),
              C_TEAL, bold=True, font_size=9.5, align=PP_ALIGN.LEFT,
              text="  06   Customer Facing Interfaces", pad=Inches(0.08))

    ifaces = [f for f in info.get("customer_interfaces", []) if f][:5]
    ifaces_str = ", ".join(ifaces) if ifaces else "Website, Email Support, Sales Representatives"
    add_textbox(slide,
                RP_X + Inches(0.15), CF_Y + Inches(0.36),
                RP_W - Inches(0.3), Inches(0.52),
                f"The structured analysis will include an assessment of {company_short} "
                f"customer-facing interfaces, including {ifaces_str}.",
                font_size=8.5, color=C_BODY_LITE)

    # ─ 5e. Protiviti branding ─────────────────────────────────
    add_textbox(slide,
                SW - Inches(1.6), SH - Inches(0.38), Inches(1.4), Inches(0.3),
                "protiviti®", font_size=9, italic=True,
                color=C_BODY_LITE, align=PP_ALIGN.RIGHT)

    # Slide number
    add_textbox(slide,
                LP_X, SH - Inches(0.38), Inches(0.5), Inches(0.3),
                "5", font_size=9, bold=True, color=C_WHITE)


# ─────────────────────────────────────────────────────────────
# HOSTING TEXT BUILDER
# ─────────────────────────────────────────────────────────────
def _build_hosting_text(info: dict) -> str:
    h   = info.get("hosting_model", "")
    spec = info.get("hosting_specify", "")
    parts = []
    if "On-premise" in h or "On-Premise" in h:
        parts.append("On-Premise")
    if "Cloud" in h:
        parts.append("Cloud")
    if "Hybrid" in h:
        parts.append("Hybrid")
    mode = " / ".join(parts) if parts else (h or "On-Premise")
    if spec:
        return f"{mode} Hosting: Personal data is stored and hosted on {spec}."
    return f"{mode} Hosting: All personal data is currently stored on {mode} infrastructure."


# ─────────────────────────────────────────────────────────────
# HELPER: Scrape website
# ─────────────────────────────────────────────────────────────
def scrape_website(url: str, max_chars: int = 5000) -> str:
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=12)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        for tag in soup(["script","style","nav","footer","header"]):
            tag.decompose()
        return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:max_chars]
    except Exception as e:
        return f"[Website scrape failed: {e}]"


# ─────────────────────────────────────────────────────────────
# HELPER: Parse Pre-Scoping Questionnaire
# ─────────────────────────────────────────────────────────────
def parse_questionnaire(docx_bytes: bytes) -> dict:
    doc    = DocxDocument(io.BytesIO(docx_bytes))
    tables = doc.tables

    def cell(ti, ri, ci):
        try:
            return tables[ti].rows[ri].cells[ci].text.strip()
        except:
            return ""

    # Employee count (table 2, row 3)
    emp_raw = cell(2, 3, 2)
    emp_count = ""
    for line in emp_raw.split("\n"):
        s = line.strip()
        if s and not s.startswith("If") and not s.startswith("Please"):
            if any(x in s for x in ["<","–",">"]):
                emp_count = s
                break
    for line in emp_raw.split("\n"):
        if "specify" in line.lower():
            val = line.split(":")[-1].strip().replace("_","").strip()
            if val:
                emp_count = val
                break
    if not emp_count:
        emp_count = emp_raw.split("\n")[0].strip()

    # Policy status (table 4, row 3)
    policy_raw = cell(4, 3, 2)
    policy_status = ""
    for line in policy_raw.split("\n"):
        s = line.strip()
        if s and "specify" in line.lower():
            val = line.split(":")[-1].strip().replace("_","").strip()
            if val:
                policy_status = val
                break
    if not policy_status:
        policy_status = policy_raw.split("\n")[0].strip()

    # Business lines (table 6, row 1)
    blines_raw = cell(6, 1, 2)
    blines = []
    for line in blines_raw.split("\n"):
        s = line.strip()
        if (s and not s.startswith("Specify") and not s.startswith("Please")
                and not s.startswith("Other") and not s.startswith("__")
                and not s.startswith("JV")):
            blines.append(s)

    # Departments (table 6, row 2)
    dept_raw = cell(6, 2, 2)
    depts = []
    extra = ""
    for line in dept_raw.split("\n"):
        s = line.strip()
        if "Departments" in s and ":" in s:
            extra = s.split(":",1)[-1].strip()
        elif (s and not s.startswith("Other") and not s.startswith("Specify")
              and not s.startswith("__") and "Departments" not in s):
            depts.append(s)
    if extra:
        for d in re.split(r"[,/]", extra):
            d = d.strip()
            if d and len(d) > 2:
                depts.append(d)
    if not depts:
        depts = ["HR & People Operations","IT & Cybersecurity","Legal & Compliance",
                 "Finance","Sales","Manufacturing Operations"]

    # Applications (table 8, row 2)
    apps_raw = cell(8, 2, 2)
    apps = []
    for line in apps_raw.split("\n"):
        s = line.strip()
        if (s and not s.startswith("Other") and not s.startswith("Specify")
                and not s.startswith("__")):
            apps.append(s)

    # Interfaces (table 8, row 1)
    iface_raw = cell(8, 1, 2)
    ifaces = []
    for line in iface_raw.split("\n"):
        s = line.strip()
        if (s and not s.startswith("Other") and not s.startswith("Specify")
                and not s.startswith("Please") and not s.startswith("__")):
            ifaces.append(s)

    # Hosting (table 8, row 4)
    hosting_raw = cell(8, 4, 2)
    hosting = hosting_raw.split("\n")[0].strip()
    hosting_spec = ""
    for line in hosting_raw.split("\n"):
        if ("ERP" in line or "SAP" in line or "HRMS" in line or
                "specify" in line.lower()):
            hosting_spec = line.split(":")[-1].strip().replace("_","").strip()
            break

    # Data subjects (table 10, row 1)
    subj_raw = cell(10, 1, 2)
    subjects = []
    for line in subj_raw.split("\n"):
        s = re.sub(r"\s*\(.*\)","", line.strip()).strip()
        if (s and not s.startswith("Other") and not s.startswith("Specify")
                and not s.startswith("Please") and not s.startswith("__")):
            subjects.append(s)

    # Data types (table 10, row 2)
    dtype_raw = cell(10, 2, 2)
    dtypes = []
    for line in dtype_raw.split("\n"):
        s = re.sub(r"\s*\(.*\)","", line.strip()).strip()
        if (s and not s.startswith("Other") and not s.startswith("Specify")
                and not s.startswith("__")):
            dtypes.append(s)

    return {
        "employee_count":       emp_count or "—",
        "policy_status":        policy_status,
        "core_business_lines":  blines,
        "departments":          [d for d in depts if d],
        "applications":         apps,
        "customer_interfaces":  ifaces,
        "hosting_model":        hosting,
        "hosting_specify":      hosting_spec,
        "data_subjects":        subjects,
        "data_types":           dtypes,
    }


# ─────────────────────────────────────────────────────────────
# GROQ HELPERS
# ─────────────────────────────────────────────────────────────
GROQ_MODEL = "llama-3.3-70b-versatile"

def groq_call(client: Groq, prompt: str, max_tokens: int = 900) -> str:
    r = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role":"user","content": prompt}],
        temperature=0.25,
        max_tokens=max_tokens,
    )
    return r.choices[0].message.content.strip()

def ctx(company_name, company_short, info, website_text) -> str:
    return (
        f"Company: {company_name} ({company_short})\n"
        f"Business lines: {', '.join(info['core_business_lines'])}\n"
        f"Departments: {', '.join(info['departments'])}\n"
        f"Employees: {info['employee_count']}\n"
        f"Hosting: {info['hosting_model']} – {info['hosting_specify']}\n"
        f"Applications: {', '.join(info['applications'])}\n"
        f"Interfaces: {', '.join(info['customer_interfaces'])}\n"
        f"Data subjects: {', '.join(info['data_subjects'])}\n"
        f"Website extract: {website_text[:1200]}"
    )


# ── SLIDE 4 PROMPTS (strict word limits to prevent text overflow) ──────────

P_DESC = """
Rewrite the paragraph below for a NEW company. Replace ALL Eveready/EIIL-specific facts \
(products, manufacturing type, distribution) with facts for the new company. \
STRICT LIMIT: 80–90 words maximum. Keep professional consulting tone.

NEW COMPANY:
{context}

ORIGINAL (do NOT copy — rewrite for new company):
Eveready Industries India Ltd. (EIIL) is a leading Indian manufacturer of portable energy \
and lighting solutions, operating through a diversified multi‑segment model spanning dry‑cell \
batteries, flashlights, consumer lighting, professional lighting and electrical accessories \
across domestic and select international markets. The company follows a predominantly B2B and \
B2B2C‑driven model, serving distributors, retailers, institutional buyers and large‑scale \
channel partners through one of India's widest FMCG‑style distribution networks, while \
maintaining limited B2C interfaces through brand engagement, after‑sales support and product \
service programs. EIIL enables end‑to‑end product development, high‑volume manufacturing, \
nationwide distribution and lifecycle management through technology‑driven quality systems, \
DSIR‑approved R&D capabilities, integrated manufacturing facilities and data‑enabled supply‑chain \
operations, ensuring safe, reliable, compliant and cost‑efficient delivery of portable power and \
lighting solutions across diverse consumer and commercial segments.

Return ONLY the rewritten paragraph. Max 90 words.
"""

P_SCOPE = """
Rewrite the sentence below for a NEW company. Replace industry-specific operations \
("lending, leasing and factoring operations") with the new company's actual operations. \
Keep all DPDPA/privacy language exactly as-is. Max 45 words.

NEW COMPANY:
{context}

ORIGINAL:
EIIL seeks support to establish a robust, end-to-end data privacy and personal data protection \
program aligned with the Digital Personal Data Protection Act, 2023 and applicable Rules, \
calibrated to its people, process and technology landscape across lending, leasing and \
factoring operations.

Return ONLY the rewritten sentence.
"""

P_BULLETS = """
Rewrite the 7 bullet points below for a NEW company. Replace EIIL-specific industry terms \
(manufacturing, R&D, supply chain, distribution, batteries, flashlights, distributors, retailers, \
logistics) with equivalent terms for the new company. Keep ALL privacy/compliance language \
exactly as-is. Max 30 words per bullet.

NEW COMPANY:
{context}

BULLETS (rewrite each, keep same order, return one per line):
1. Conduct an enterprise-wide applicability assessment and privacy gap analysis, covering data discovery, lifecycle mapping, inventories, RoPA and documentation of internal/external data flows across EIIL's manufacturing, R&D, supply chain, procurement, commercial, HR, enterprise systems and distribution operations.
2. Assess privacy, information security and regulatory risks across EIIL's manufacturing, quality, logistics, commercial, enterprise and SaaS platforms, including analytics environments, physical repositories and third-party networks such as distributors, retailers, logistics partners and service vendors.
3. Evaluate governance structures, policies and controls covering lawful purpose, consent (where applicable), retention, erasure, grievance handling, DPR workflows, cross-border transfers and personal data breach processes.
4. Design and operationalize a scalable privacy governance and risk framework, defining roles, accountability, escalation paths and procedures for DPIAs and risk-based reviews of new systems, digital initiatives and operational programs.
5. Support rollout of updated privacy policies, notices and procedures for consent, DPR, retention/deletion, breach response and DPIA processes, tailored for corporate, manufacturing, R&D, commercial and customer-facing teams.
6. Coordinate remediation across key platforms to strengthen consent workflows, DPR handling, third-party data sharing controls, data minimization and privacy-by-design requirements with support from selected tooling partners.
7. Deliver role-based privacy training, define governance KPIs and RACI structures and enable reporting and dashboards to support continuous oversight, audit readiness, regulatory preparedness and executive visibility.

Return numbered lines (1. text, 2. text …). No extra commentary.
"""

P_S11 = """
Rewrite the paragraph below for a NEW company. Replace "Eveready Industries India Ltd." and \
"manufacturing, supply-chain, commercial, distribution" with the new company's operations. \
Keep the privacy methodology framing. Max 50 words.

NEW COMPANY:
{context}

ORIGINAL:
For this engagement, the privacy compliance model will be applied exclusively to the internal \
functions, processes and governance structures of Eveready Industries India Ltd., supporting its \
manufacturing, supply-chain, commercial, distribution and corporate operations, which primarily \
operate through B2B and B2B2C channels, with limited B2C personal data processing through \
customer service interactions, digital platform usage and product service requests.

Return ONLY the rewritten paragraph.
"""

P_S17 = """
Rewrite all 6 Data Lifecycle paragraphs for a NEW company. Replace EVERY Eveready/EIIL-specific \
product, process and system reference with equivalent references for the new company. \
Keep privacy/data-governance framing identical. Max 55 words per section.

NEW COMPANY:
{context}

Return a JSON object with keys: \
"collection","use_processing","storage","sharing","retention","disposal"
Each value = the rewritten paragraph text only (no title prefix).
Return ONLY valid JSON, no markdown fences.

ORIGINALS:
collection: We will review how Eveready Industries India Ltd. (EIIL) collects personal, operational and regulatory data across functions such as employee onboarding; manufacturing processes for batteries, flashlights and lighting products; distributor onboarding; sales operations; supply-chain coordination; and customer service requirements. This includes data captured through ERP systems, plant-level manufacturing platforms, distributor management portals, helpdesk and CRM interfaces.
use_processing: We will assess how collected data is used for manufacturing planning, quality control, inventory management, supply chain coordination, compliance reporting and performance monitoring across EIIL's key segments: batteries, flashlights, consumer lighting, professional lighting and electrical accessories. This includes data integration across systems such as ERP, CRM, distributor management systems and quality monitoring platforms.
storage: We will examine secure storage of manufacturing, safety, employee and vendor data across cloud platforms, on-premise servers at EIIL's manufacturing units, validated production systems, backup systems and R&D repositories. Controls for authentication, role-based access, audit trails and compliance with applicable industry and corporate guidelines will also be reviewed.
sharing: We will evaluate data-sharing practices with distributors, logistics partners, manufacturing vendors, regulatory authorities, retailers and internal teams. This includes reviewing contractual safeguards, supply-chain data-processing requirements, cross-border data transfer practices (where applicable), anonymization procedures and security measures.
retention: We will review retention policies for manufacturing logs, quality-control reports, product testing data, R&D records, HR and payroll files, vendor documentation, distributor agreements, operational logs and financial documentation. Retention requirements will be assessed against regulatory mandates, audit requirements and internal EIIL governance policies.
disposal: We will verify secure deletion, destruction and anonymization of records across digital platforms, manufacturing systems, archival repositories, distributor management systems and physical documentation. Disposal workflows will be reviewed for alignment with regulatory expectations and internal EIIL data-governance guidelines.
"""


# ─────────────────────────────────────────────────────────────
# PPTX TEXT REPLACEMENT (all other slides)
# ─────────────────────────────────────────────────────────────
def _rep_para(para, rep: dict):
    full = "".join(r.text for r in para.runs)
    if not any(k in full for k in rep if k):
        return
    new = full
    for k, v in rep.items():
        if k and v is not None:
            new = new.replace(k, v)
    if new != full and para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ""

def _rep_shape(shape, rep: dict):
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            _rep_para(p, rep)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for p in cell.text_frame.paragraphs:
                    _rep_para(p, rep)
    if shape.shape_type == 6:
        for s in shape.shapes:
            _rep_shape(s, rep)

def rep_slide(slide, rep: dict):
    for shape in slide.shapes:
        _rep_shape(shape, rep)

def set_para_text(slide, shape_name: str, fragment: str, new_text: str) -> bool:
    for shape in slide.shapes:
        if shape_name and shape.name != shape_name:
            continue
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full = "".join(r.text for r in para.runs)
            if fragment in full:
                if para.runs:
                    para.runs[0].text = new_text
                    for r in para.runs[1:]:
                        r.text = ""
                return True
    return False


# ─────────────────────────────────────────────────────────────
# MAIN BUILD
# ─────────────────────────────────────────────────────────────
def build_presentation(pptx_bytes, company_name, company_short, info, ai):
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Global name replacements
    gmap = {
        "Eveready Industries India Ltd. (EIIL)": company_name,
        "Eveready Industries India Ltd":         company_name.split("(")[0].strip(),
        "Eveready Industries":                   company_name.split("(")[0].strip(),
        " EIIL'":  f" {company_short}'",
        " EIIL's": f" {company_short}'s",
        "(EIIL)":  f"({company_short})",
        "EIIL":    company_short,
    }
    for slide in prs.slides:
        rep_slide(slide, gmap)

    # ── Slide 4 (index 3) ────────────────────────────────────
    if len(prs.slides) > 3:
        s4 = prs.slides[3]
        set_para_text(s4, "TextBox 8", "leading", ai["s4_desc"])
        set_para_text(s4, "TextBox 3", "seeks support", ai["s4_scope"])
        for frag, new_b in zip(
            ["Conduct an enterprise","Assess privacy, information security",
             "Evaluate governance","Design and operationalize",
             "Support rollout","Coordinate remediation","Deliver role"],
            ai.get("s4_bullets", [])
        ):
            set_para_text(s4, "TextBox 3", frag, new_b)

    # ── Slide 5 (index 4) — full rebuild ─────────────────────
    if len(prs.slides) > 4:
        rebuild_slide5(prs.slides[4], company_name, company_short, info)

    # ── Slide 11 (index 10) ──────────────────────────────────
    if len(prs.slides) > 10:
        set_para_text(prs.slides[10], "Rectangle 1",
                      "For this engagement", ai["s11"])

    # ── Slide 17 (index 16) ──────────────────────────────────
    if len(prs.slides) > 16:
        s17  = prs.slides[16]
        lc   = ai.get("s17_lifecycle", {})
        for shp_name, key in [
            ("Rectangle 41","collection"), ("Rectangle 8","use_processing"),
            ("Rectangle 9","storage"),     ("Rectangle 10","sharing"),
            ("Rectangle 11","retention"),  ("Rectangle 12","disposal"),
        ]:
            if lc.get(key):
                set_para_text(s17, shp_name, "We will", lc[key])

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# GENERATE FLOW
# ─────────────────────────────────────────────────────────────
if generate_btn:
    errs = []
    if not groq_api_key:    errs.append("🔑 Groq API Key required.")
    if not company_name:    errs.append("🏢 Company name required.")
    if not company_short:   errs.append("🏷️ Abbreviation required.")
    if not company_website: errs.append("🌐 Website URL required.")
    if not uploaded_ppt:    errs.append("📁 Template PPTX required.")
    if not uploaded_docx:   errs.append("📝 Questionnaire .docx required.")
    for e in errs: st.error(e)
    if errs: st.stop()

    pptx_bytes = uploaded_ppt.read()
    docx_bytes = uploaded_docx.read()
    client     = Groq(api_key=groq_api_key)

    # Step 1 – parse questionnaire
    with st.status("📋 Parsing questionnaire…", expanded=True) as s:
        info = parse_questionnaire(docx_bytes)
        s.update(label="✅ Questionnaire parsed", state="complete")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📋 Extracted from Questionnaire")
        st.json({k: v for k, v in info.items() if v not in (None,"",[],"—")})

    # Step 2 – scrape
    with st.status("🌐 Scraping website…", expanded=False) as s:
        web = scrape_website(company_website)
        ok  = not web.startswith("[Website")
        s.update(label="✅ Website scraped" if ok
                 else "⚠️ Scrape failed — using questionnaire data only",
                 state="complete" if ok else "error")

    context = ctx(company_name, company_short, info, web)

    # Step 3 – AI
    ai = {}
    with st.status("🤖 Generating AI content…", expanded=True) as s:
        def safe_call(key, prompt, max_tok=900, fallback=""):
            st.write(f"  → {key}…")
            try:
                ai[key] = groq_call(client, prompt, max_tok)
            except Exception as e:
                ai[key] = fallback
                st.warning(f"{key}: {e}")

        safe_call("s4_desc",    P_DESC.format(context=context),    700,  f"{company_name} is a leading specialty packaging company…")
        safe_call("s4_scope",   P_SCOPE.format(context=context),   300,  f"{company_short} seeks support to establish a robust DPDPA privacy program…")
        raw_b = ""
        try:
            st.write("  → s4_bullets…")
            raw_b = groq_call(client, P_BULLETS.format(context=context), 1400)
            lines = [re.sub(r"^\d+\.\s*","", l).strip()
                     for l in raw_b.split("\n") if l.strip()]
            ai["s4_bullets"] = lines
        except Exception as e:
            ai["s4_bullets"] = []
            st.warning(f"s4_bullets: {e}")

        safe_call("s11",        P_S11.format(context=context),     400,
                  f"For this engagement, the privacy compliance model will be applied to the internal functions and governance structures of {company_name}, supporting its {', '.join(info['core_business_lines'][:2])} and corporate operations.")
        try:
            st.write("  → s17_lifecycle…")
            raw17 = groq_call(client, P_S17.format(context=context), 2200)
            raw17 = re.sub(r"^```(?:json)?","", raw17).strip()
            raw17 = re.sub(r"```$","", raw17).strip()
            ai["s17_lifecycle"] = json.loads(raw17)
        except Exception as e:
            ai["s17_lifecycle"] = {}
            st.warning(f"s17_lifecycle: {e}")

        s.update(label="✅ All AI content generated", state="complete")

    # Step 4 – build PPTX
    with st.status("📝 Building PPTX…", expanded=False) as s:
        try:
            output = build_presentation(pptx_bytes, company_name, company_short, info, ai)
            s.update(label="✅ PPTX ready!", state="complete")
        except Exception as e:
            s.update(label="❌ Failed", state="error")
            st.error(str(e))
            import traceback; st.code(traceback.format_exc())
            st.stop()

    st.success("🎉 Proposal generated!")

    with st.expander("🔍 Preview AI content"):
        st.markdown("**Slide 4 – Company Description:**"); st.info(ai.get("s4_desc",""))
        st.markdown("**Slide 4 – Scope Paragraph:**");     st.info(ai.get("s4_scope",""))
        if ai.get("s4_bullets"):
            st.markdown("**Slide 4 – Bullets:**")
            for b in ai["s4_bullets"]: st.write(f"• {b}")
        st.markdown("**Slide 11 – Operating Model:**");    st.info(ai.get("s11",""))
        if ai.get("s17_lifecycle"):
            st.markdown("**Slide 17 – Data Lifecycle:**")
            for k, v in ai["s17_lifecycle"].items():
                st.write(f"**{k.replace('_',' ').title()}:** {v}")

    safe  = re.sub(r"[^\w\s-]","", company_name).strip().replace(" ","_")[:40]
    fname = f"Proposal_Data_Privacy_{safe}_March2026.pptx"
    st.download_button("⬇️ Download Personalised Proposal",
                       output, fname,
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                       use_container_width=True, type="primary")
    st.caption("ℹ️ Fields without sufficient data are left as in the template for human review.")

else:
    st.info("👈 Fill in all details in the sidebar, upload both files, then click **Generate Proposal**.")
    with st.expander("📖 How it works"):
        st.markdown("""
| Slide | Change | Source |
|---|---|---|
| **1** | Company name | Input |
| **4** | Company description + scope + bullets | AI (Groq, word-limited) |
| **5** | **Full professional redesign** — employee count, hosting, apps, departments, data types & subjects | Questionnaire |
| **11** | Operating model paragraph | AI |
| **12, 14, 19** | Name auto-replace | Auto |
| **17** | All 6 Data Lifecycle sections | AI |

All design, colours, fonts, charts, team slides and Protiviti content remain untouched.
""")
