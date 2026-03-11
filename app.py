"""
AI-Powered Proposal Generator
Personalises a DPDPA consulting proposal template for any target company.

Slides modified:
  Slide 1  – Company name + date
  Slide 4  – Company description paragraph + scope understanding (AI-written)
  Slide 5  – Employee count, hosting, applications, departments, data subjects (from docx)
  Slide 11 – Privacy Operating Model paragraph (AI-written)
  Slide 12 – Phase bullets: EIIL references replaced
  Slide 14 – Phase I: EIIL references replaced
  Slide 17 – Data Lifecycle 6 paragraphs (AI-written per section)
  Slide 19 – Phase II: EIIL references replaced
"""

import io, json, re, copy
import streamlit as st
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Pt
from docx import Document as DocxDocument
from groq import Groq

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Proposal Generator – Protiviti",
    page_icon="📊",
    layout="wide",
)

st.title("📊 AI Proposal Generator")
st.markdown(
    "Upload the **template PPTX**, the **filled Pre-Scoping Questionnaire (.docx)**, "
    "enter basic company details and your **Groq API key** — the app will personalise "
    "every relevant slide while keeping design, fonts and layout 100% intact."
)

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key = st.text_input("Groq API Key", type="password", placeholder="gsk_...")
    st.markdown("---")
    st.header("🏢 Company Details")
    company_name   = st.text_input("Full Company Name", placeholder="SGD Pharma India Pvt. Ltd.")
    company_short  = st.text_input("Abbreviation / Short Name", placeholder="SGD")
    company_website = st.text_input("Website URL", placeholder="https://www.sgd-pharma.com")
    st.markdown("---")
    st.header("📁 Files")
    uploaded_ppt  = st.file_uploader("Template PPTX", type=["pptx"])
    uploaded_docx = st.file_uploader("Pre-Scoping Questionnaire (.docx)", type=["docx"])
    st.markdown("---")
    generate_btn = st.button("🚀 Generate Proposal", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────
# HELPER: Scrape website
# ─────────────────────────────────────────────────────────────
def scrape_website(url: str, max_chars: int = 6000) -> str:
    try:
        resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=12)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        for tag in soup(["script","style","nav","footer","header"]):
            tag.decompose()
        text = re.sub(r"\s+", " ", soup.get_text(separator=" ", strip=True))
        return text[:max_chars]
    except Exception as e:
        return f"[Website scrape failed: {e}]"

# ─────────────────────────────────────────────────────────────
# HELPER: Parse Pre-Scoping Questionnaire
# ─────────────────────────────────────────────────────────────
def parse_questionnaire(docx_bytes: bytes) -> dict:
    """
    Reads the structured questionnaire tables and returns a dict of answers.
    The questionnaire has alternating header tables + data tables.
    Table indices (0-indexed):
      0  – title row
      1  – ORGANISATIONAL OVERVIEW header
      2  – org data  (rows: subsidiaries, centralised IT/HR/Legal, employee strength)
      3  – GOVERNANCE header
      4  – governance data  (privacy committee, decision makers, policy status)
      5  – BUSINESS LINES header
      6  – business data  (core lines, key stakeholders)
      7  – DATA ECOSYSTEM header
      8  – ecosystem data  (customer interfaces, apps, data discovery, hosting)
      9  – DATA SUBJECTS header
      10 – data subjects data  (data subject categories, data types)
      11 – footer
    """
    doc = DocxDocument(io.BytesIO(docx_bytes))
    tables = doc.tables

    def cell(ti, ri, ci):
        try:
            return tables[ti].rows[ri].cells[ci].text.strip()
        except:
            return ""

    def pick_checked(text: str) -> list:
        """Return lines that look 'selected' (no leading spaces = checked in this doc)."""
        lines = text.split("\n")
        # Lines without leading whitespace are the checked items (template uses indent for unchecked)
        checked = [l.strip() for l in lines if l and not l.startswith("  ")]
        return [c for c in checked if c and not c.lower().startswith("if ") 
                and not c.lower().startswith("please") and c != "Yes" and c != "No"]

    def first_line(text: str) -> str:
        return text.split("\n")[0].strip()

    # ── Org Overview (table index 2) ──────────────────────────
    subsidiaries_text  = cell(2, 1, 2)
    emp_strength_text  = cell(2, 3, 2)

    # Extract subsidiary info
    has_subsidiaries = "Yes" in subsidiaries_text.split("\n")[0] or \
                       not subsidiaries_text.startswith("  ")
    subsidiary_detail = ""
    for line in subsidiaries_text.split("\n"):
        line = line.strip()
        if line and not line in ("Yes","No") and not line.startswith("If Yes"):
            subsidiary_detail = line
            break

    # Extract employee count
    emp_count = ""
    for line in emp_strength_text.split("\n"):
        line = line.strip()
        if line and line not in ("< 500","500 – 1,000","1,000 – 5,000","> 5,000") \
                and not line.startswith("If > 5,000"):
            if ">" in line or "<" in line or "–" in line:
                emp_count = line
                break
    # Also check for explicit spec
    for line in emp_strength_text.split("\n"):
        if "specify" in line.lower():
            val = line.split(":")[-1].strip().replace("__","").strip()
            if val:
                emp_count = val
                break
    if not emp_count:
        emp_count = first_line(emp_strength_text)

    # ── Governance (table index 4) ──────────────────────────
    policy_status_text = cell(4, 3, 2)
    policy_status = ""
    for line in policy_status_text.split("\n"):
        line = line.strip()
        if line and line not in ("Yes, centralised global office","Yes, regional offices",
                                  "No, decisions taken by IT / Legal / Other","No formal structure",
                                  "Existing framework in place (requires update)",
                                  "Drafted but not implemented",
                                  "Needs to be formulated from scratch","Other"):
            policy_status = line
            break
    if not policy_status:
        # pick first non-empty, non-option line
        for line in policy_status_text.split("\n"):
            s = line.strip()
            if s:
                policy_status = s
                break

    # ── Business Lines (table index 6) ──────────────────────
    core_lines_text   = cell(6, 1, 2)
    stakeholders_text = cell(6, 2, 2)

    core_lines = []
    for line in core_lines_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Specify") and not line.startswith("Please") \
                and not line.startswith("Other") and not line.startswith("__"):
            core_lines.append(line)

    departments = []
    for line in stakeholders_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Other") and not line.startswith("Specify") \
                and not line.startswith("__") and "Departments" not in line:
            # pull out department name from lines like "Departments - Finance, Purchase..."
            if "Departments" in line or "," in line:
                for dept in re.split(r"[,\-]", line):
                    dept = dept.strip()
                    if dept and len(dept) > 2:
                        departments.append(dept)
            else:
                departments.append(line)
    if not departments:
        departments = ["HR & People Operations","IT & Cybersecurity","Legal & Compliance",
                       "Finance","Sales","Manufacturing Operations"]

    # ── Data Ecosystem (table index 8) ──────────────────────
    interfaces_text   = cell(8, 1, 2)
    apps_text         = cell(8, 2, 2)
    hosting_text      = cell(8, 4, 2)

    interfaces = []
    for line in interfaces_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Other") and not line.startswith("Specify") \
                and not line.startswith("Please") and not line.startswith("__"):
            interfaces.append(line)

    apps = []
    for line in apps_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Other") and not line.startswith("Specify") \
                and not line.startswith("__"):
            apps.append(line)

    # Hosting
    hosting = ""
    for line in hosting_text.split("\n"):
        s = line.strip()
        if s and not s.startswith("__"):
            hosting = s
            break
    hosting_specify = ""
    for line in hosting_text.split("\n"):
        if "specify" in line.lower() or "ERP" in line or "SAP" in line or "HRMS" in line:
            val = line.split(":")[-1].strip().replace("_","").strip()
            if val:
                hosting_specify = val
                break

    # ── Data Subjects (table index 10) ──────────────────────
    subjects_text   = cell(10, 1, 2)
    data_types_text = cell(10, 2, 2)

    data_subjects = []
    for line in subjects_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Other") and not line.startswith("Specify") \
                and not line.startswith("Please") and not line.startswith("__"):
            # Clean trailing parenthetical
            name = re.sub(r"\s*\(.*\)", "", line).strip()
            if name:
                data_subjects.append(name)

    data_types = []
    for line in data_types_text.split("\n"):
        line = line.strip()
        if line and not line.startswith("Other") and not line.startswith("Specify") \
                and not line.startswith("__"):
            name = re.sub(r"\s*\(.*\)", "", line).strip()
            if name:
                data_types.append(name)

    return {
        "has_subsidiaries": has_subsidiaries,
        "subsidiary_detail": subsidiary_detail,
        "employee_count": emp_count or "—",
        "policy_status": policy_status,
        "core_business_lines": core_lines,
        "departments": [d for d in departments if d],
        "customer_interfaces": interfaces,
        "applications": apps,
        "hosting_model": hosting,
        "hosting_specify": hosting_specify,
        "data_subjects": data_subjects,
        "data_types": data_types,
    }

# ─────────────────────────────────────────────────────────────
# GROQ HELPERS
# ─────────────────────────────────────────────────────────────
GROQ_MODEL = "llama-3.3-70b-versatile"

def groq_call(client: Groq, prompt: str, max_tokens: int = 1500) -> str:
    resp = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.25,
        max_tokens=max_tokens,
    )
    return resp.choices[0].message.content.strip()


def build_context(company_name, company_short, info: dict, website_text: str) -> str:
    return f"""
Company Full Name: {company_name}
Company Short Name: {company_short}
Industry: Specialty Glass & Pharmaceutical Packaging
Business Lines: {", ".join(info["core_business_lines"])}
Departments in scope: {", ".join(info["departments"])}
Employee Count: {info["employee_count"]}
Hosting: {info["hosting_model"]} — {info["hosting_specify"]}
Applications: {", ".join(info["applications"])}
Customer-facing interfaces: {", ".join(info["customer_interfaces"])}
Data Subjects: {", ".join(info["data_subjects"])}
Data Types: {", ".join(info["data_types"])}
Website extract: {website_text[:1500]}
""".strip()


# ── Slide 4: Company description paragraph ──────────────────
SLIDE4_DESC_PROMPT = """
You are writing content for a professional data-privacy consulting proposal.

Rewrite the paragraph below — originally written for Eveready Industries India Ltd. (EIIL) — 
so it accurately describes the NEW company. Replace every industry-specific fact 
(products, manufacturing type, distribution model, capabilities) with facts relevant to 
the new company. Keep the sentence structure, professional tone, and approximate length.

NEW COMPANY CONTEXT:
{context}

ORIGINAL PARAGRAPH:
Eveready Industries India Ltd. (EIIL) is a leading Indian manufacturer of portable energy and \
lighting solutions, operating through a diversified multi‑segment model spanning dry‑cell batteries, \
flashlights, consumer lighting, professional lighting and electrical accessories across domestic and \
select international markets. The company follows a predominantly B2B and B2B2C‑driven model, serving \
distributors, retailers, institutional buyers and large‑scale channel partners through one of India's \
widest FMCG‑style distribution networks, while maintaining limited B2C interfaces through brand \
engagement, after‑sales support and product service programs. EIIL enables end‑to‑end product \
development, high‑volume manufacturing, nationwide distribution and lifecycle management through \
technology‑driven quality systems, DSIR‑approved R&D capabilities, integrated manufacturing \
facilities and data‑enabled supply‑chain operations, ensuring safe, reliable, compliant and \
cost‑efficient delivery of portable power and lighting solutions across diverse consumer and commercial segments.

Return ONLY the rewritten paragraph. No labels, no quotes, no extra text.
"""

# ── Slide 4: Scope understanding paragraph ──────────────────
SLIDE4_SCOPE_PROMPT = """
Rewrite the scope paragraph below, replacing ALL references to EIIL / Eveready Industries / \
"lending, leasing and factoring operations" with the new company's actual operations.
Adjust any industry-specific wording (manufacturing, distribution channels, supply chain) 
to match the new company. Keep compliance/privacy content exactly as-is. Same length.

NEW COMPANY CONTEXT:
{context}

ORIGINAL:
EIIL seeks support to establish a robust, end-to-end data privacy and personal data protection \
program aligned with the Digital Personal Data Protection Act, 2023 and applicable Rules, \
calibrated to its people, process and technology landscape across lending, leasing and factoring operations.

Return ONLY the rewritten paragraph.
"""

# ── Slide 4: Bullet points (multiple, one per original bullet) ──
SLIDE4_BULLETS_PROMPT = """
Rewrite each bullet point below, replacing EIIL-specific industry terms \
(manufacturing, R&D, supply chain, distribution, batteries, flashlights, lighting, distributors, \
retailers, logistics partners) with terminology matching the new company.

Keep ALL privacy/compliance language exactly as-is. Keep each bullet approximately the same length.
Return ONLY the rewritten bullets, one per line, in the SAME ORDER.

NEW COMPANY CONTEXT:
{context}

ORIGINAL BULLETS (one per line):
{bullets}
"""

# ── Slide 11: Privacy Operating Model paragraph ──────────────
SLIDE11_PROMPT = """
Rewrite the paragraph below, replacing EIIL-specific language with language 
matching the new company's industry and operations. 
Keep the privacy-methodology content exactly as-is.

NEW COMPANY CONTEXT:
{context}

ORIGINAL:
For this engagement, the privacy compliance model will be applied exclusively to the internal \
functions, processes and governance structures of Eveready Industries India Ltd., supporting its \
manufacturing, supply‑chain, commercial, distribution and corporate operations, which primarily \
operate through B2B and B2B2C channels, with limited B2C personal data processing through \
customer service interactions, digital platform usage and product service requests.

Return ONLY the rewritten paragraph (single paragraph, no labels).
"""

# ── Slide 17: Data Lifecycle (6 sections) ──────────────────
SLIDE17_PROMPT = """
Rewrite all six Data Lifecycle section paragraphs below for the new company.
Replace EVERY reference to Eveready/EIIL-specific products, operations, systems and processes \
with equivalent references for the new company. Keep the privacy/data-governance framing identical.

NEW COMPANY CONTEXT:
{context}

Return a JSON object with exactly these keys: 
"collection", "use_processing", "storage", "sharing", "retention", "disposal"
Each value = the rewritten paragraph (no section title, just the text).
Return ONLY valid JSON, no markdown fences.

ORIGINALS:
01. Data Collection: We will review how Eveready Industries India Ltd. (EIIL) collects personal, \
operational and regulatory data across functions such as employee onboarding; manufacturing processes \
for batteries, flashlights and lighting products; distributor onboarding; sales operations; \
supply‑chain coordination; and customer service requirements. This includes data captured through \
ERP systems, plant‑level manufacturing platforms, distributor management portals, helpdesk \
and CRM interfaces.

02. Data Use & Processing: We will assess how collected data is used for manufacturing planning, \
quality control, inventory management, supply chain coordination, compliance reporting and \
performance monitoring across EIIL's key segments: batteries, flashlights, consumer lighting, \
professional lighting and electrical accessories. This includes data integration across systems \
such as ERP, CRM, distributor management systems and quality monitoring platforms.

03. Data Storage: We will examine secure storage of manufacturing, safety, employee and vendor data \
across cloud platforms, on‑premise servers at EIIL's manufacturing units, validated production systems, \
backup systems and R&D repositories. Controls for authentication, role‑based access, audit trails and \
compliance with applicable industry and corporate guidelines will also be reviewed.

04. Data Sharing: We will evaluate data‑sharing practices with distributors, logistics partners, \
manufacturing vendors, regulatory authorities, retailers and internal teams. This includes reviewing \
contractual safeguards, supply‑chain data‑processing requirements, cross‑border data transfer \
practices (where applicable), anonymization procedures and security measures.

05. Data Retention: We will review retention policies for manufacturing logs, quality‑control reports, \
product testing data, R&D records, HR and payroll files, vendor documentation, distributor agreements, \
operational logs and financial documentation. Retention requirements will be assessed against regulatory \
mandates, audit requirements and internal EIIL governance policies.

06. Data Disposal: We will verify secure deletion, destruction and anonymization of records across \
digital platforms, manufacturing systems, archival repositories, distributor management systems and \
physical documentation. Disposal workflows will be reviewed for alignment with regulatory expectations \
and internal EIIL data‑governance guidelines to ensure safe and compliant handling of obsolete data.
"""

# ─────────────────────────────────────────────────────────────
# PPTX MANIPULATION HELPERS
# ─────────────────────────────────────────────────────────────
def replace_in_para(para, replacements: dict):
    """Replace across all runs of a paragraph preserving formatting of run[0]."""
    full = "".join(r.text for r in para.runs)
    if not any(k in full for k in replacements if k):
        return
    new = full
    for old, nw in replacements.items():
        if old and nw is not None:
            new = new.replace(old, nw)
    if new != full and para.runs:
        para.runs[0].text = new
        for r in para.runs[1:]:
            r.text = ""


def replace_in_shape(shape, replacements: dict):
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            replace_in_para(para, replacements)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    replace_in_para(para, replacements)
    if shape.shape_type == 6:
        for s in shape.shapes:
            replace_in_shape(s, replacements)


def replace_slide(slide, replacements: dict):
    for shape in slide.shapes:
        replace_in_shape(shape, replacements)


def set_paragraph_text(slide, shape_name: str, old_fragment: str,
                        new_text: str, occurrence: int = 0) -> bool:
    """
    Find the paragraph in shape_name whose combined text contains old_fragment
    (n-th occurrence) and replace its full text with new_text.
    """
    found = 0
    for shape in slide.shapes:
        if shape_name and shape.name != shape_name:
            continue
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full = "".join(r.text for r in para.runs)
            if old_fragment in full:
                if found == occurrence:
                    if para.runs:
                        para.runs[0].text = new_text
                        for r in para.runs[1:]:
                            r.text = ""
                    return True
                found += 1
    return False


def set_textbox_full(slide, shape_name: str, new_text: str) -> bool:
    """Replace entire text content of a named shape (first paragraph)."""
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            if shape.text_frame.paragraphs:
                para = shape.text_frame.paragraphs[0]
                if para.runs:
                    para.runs[0].text = new_text
                    for r in para.runs[1:]:
                        r.text = ""
            return True
    return False


def get_shape_full_text(slide, shape_name: str) -> str:
    for shape in slide.shapes:
        if shape.name == shape_name and shape.has_text_frame:
            return "\n".join(
                "".join(r.text for r in p.runs)
                for p in shape.text_frame.paragraphs
            )
    return ""

# ─────────────────────────────────────────────────────────────
# HOSTING TEXT BUILDER
# ─────────────────────────────────────────────────────────────
def build_hosting_text(info: dict) -> str:
    h = info["hosting_model"]
    spec = info["hosting_specify"]
    parts = []
    if "On-premise" in h:
        parts.append("On-Premise")
    if "Cloud" in h:
        parts.append("Cloud")
    if "Hybrid" in h:
        parts.append("Hybrid")
    mode = "/".join(parts) if parts else h
    base = f"{mode} Hosting"
    if spec:
        return f"{base}: Personal data is stored and hosted on {spec}."
    return f"{base}: Personal data is currently stored on {mode} infrastructure."

# ─────────────────────────────────────────────────────────────
# MAIN PPTX BUILDER
# ─────────────────────────────────────────────────────────────
def build_presentation(
    pptx_bytes: bytes,
    company_name: str,
    company_short: str,
    info: dict,
    ai: dict,
) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))

    # Global name replacements applied to EVERY slide
    global_rep = {
        "Eveready Industries India Ltd. (EIIL)": company_name,
        "Eveready Industries India Ltd": company_name.split("(")[0].strip().rstrip(),
        "Eveready Industries": company_name.split("(")[0].strip().rstrip(),
        " EIIL'": f" {company_short}'",
        " EIIL's": f" {company_short}'s",
        "(EIIL)": f"({company_short})",
        "EIIL": company_short,
    }
    for slide in prs.slides:
        replace_slide(slide, global_rep)

    # ── SLIDE 1 ──────────────────────────────────────────────
    # Already handled by global replacements (company name + date is unchanged = March 2026)

    # ── SLIDE 4 (index 3) ───────────────────────────────────
    if len(prs.slides) > 3:
        s4 = prs.slides[3]
        # Company description paragraph (TextBox 8)
        set_paragraph_text(s4, "TextBox 8", "leading Indian manufacturer", ai["slide4_desc"])
        # Scope understanding paragraph (TextBox 3, first paragraph)
        set_paragraph_text(s4, "TextBox 3", "seeks support to establish", ai["slide4_scope"])
        # Bullet paragraphs in TextBox 3 (index 1 onward)
        bullets = ai.get("slide4_bullets", [])
        bullet_originals = [
            "Conduct an enterprise",
            "Assess privacy, information security",
            "Evaluate governance structures",
            "Design and operationalize",
            "Support rollout of updated privacy policies",
            "Coordinate remediation across key platforms",
            "Deliver role",
        ]
        for i, (frag, new_bullet) in enumerate(zip(bullet_originals, bullets)):
            set_paragraph_text(s4, "TextBox 3", frag, new_bullet)

    # ── SLIDE 5 (index 4) ───────────────────────────────────
    if len(prs.slides) > 4:
        s5 = prs.slides[4]

        # Employee count
        replace_slide(s5, {"1200+": info["employee_count"]})

        # Hosting text
        hosting_txt = build_hosting_text(info)
        set_paragraph_text(s5, "TextBox 31",
            "100% On-Premise Hosting", hosting_txt)

        # Application Ecosystem
        apps_str = ", ".join(info["applications"][:6]) if info["applications"] else "ERP, CRM, HRMS"
        set_paragraph_text(s5, "TextBox 19",
            "utilizes core enterprise applications",
            f"{company_short} utilizes core enterprise applications including {apps_str}.")

        # Departments
        depts = info["departments"]
        n = len(depts)
        dept_str = ", ".join(depts)
        set_paragraph_text(s5, "TextBox 38",
            "Audit coverage includes assessment",
            f"Audit coverage includes assessment of data handling practices within "
            f"{n} departments such as {dept_str}")

        # Data Subjects (individual text boxes)
        subject_shapes = ["TextBox 102","TextBox 104","TextBox 106",
                          "TextBox 108","TextBox 111","TextBox 113"]
        subjects = info["data_subjects"]
        for i, shape_name in enumerate(subject_shapes):
            if i < len(subjects):
                set_textbox_full(s5, shape_name, subjects[i])
            else:
                set_textbox_full(s5, shape_name, "")

        # Critical data types label row (TextBox 96 = "Service logs & complaints")
        if info["data_types"]:
            set_textbox_full(s5, "TextBox 96", info["data_types"][-1] if len(info["data_types"]) > 1 else "")

    # ── SLIDE 11 (index 10) ──────────────────────────────────
    if len(prs.slides) > 10:
        s11 = prs.slides[10]
        set_paragraph_text(s11, "Rectangle 1",
            "For this engagement, the privacy compliance model",
            ai["slide11_para"])

    # ── SLIDE 12 (index 11) ──────────────────────────────────
    # Global replacements already handled; no extra AI content needed here.

    # ── SLIDE 14 (index 13) ──────────────────────────────────
    # Global replacements already handled.

    # ── SLIDE 17 (index 16) ──────────────────────────────────
    if len(prs.slides) > 16:
        s17 = prs.slides[16]
        lc = ai.get("slide17_lifecycle", {})
        mapping = {
            "Rectangle 41": ("01. Data Collection", lc.get("collection","")),
            "Rectangle 8":  ("02. Data Use & Processing", lc.get("use_processing","")),
            "Rectangle 9":  ("03. Data Storage", lc.get("storage","")),
            "Rectangle 10": ("04. Data Sharing", lc.get("sharing","")),
            "Rectangle 11": ("05. Data Retention", lc.get("retention","")),
            "Rectangle 12": ("06. Data Disposal", lc.get("disposal","")),
        }
        for shape_name, (title, new_para) in mapping.items():
            if new_para:
                # Shape has two paras: title + text. We target the text para.
                set_paragraph_text(s17, shape_name, "We will", new_para)

    # ── SLIDE 19 (index 18) ──────────────────────────────────
    # Global replacements already handled (EIIL → company_short in Phase II bullets)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────
# GENERATE FLOW
# ─────────────────────────────────────────────────────────────
if generate_btn:
    errors = []
    if not groq_api_key:   errors.append("🔑 Groq API Key required.")
    if not company_name:   errors.append("🏢 Company name required.")
    if not company_short:  errors.append("🏷️ Short name / abbreviation required.")
    if not company_website:errors.append("🌐 Website URL required.")
    if not uploaded_ppt:   errors.append("📁 Template PPTX required.")
    if not uploaded_docx:  errors.append("📝 Pre-Scoping Questionnaire (.docx) required.")
    for e in errors:
        st.error(e)
    if errors:
        st.stop()

    pptx_bytes = uploaded_ppt.read()
    docx_bytes = uploaded_docx.read()
    client = Groq(api_key=groq_api_key)

    # ── Step 1: Parse questionnaire ──────────────────────────
    with st.status("📋 Parsing questionnaire...", expanded=True) as status:
        info = parse_questionnaire(docx_bytes)
        status.update(label="✅ Questionnaire parsed", state="complete")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📋 Extracted from Questionnaire")
        st.json({k: v for k, v in info.items() if v not in (None, "", [], {})})

    # ── Step 2: Website scrape ───────────────────────────────
    with st.status("🌐 Scraping company website...", expanded=False) as status:
        website_text = scrape_website(company_website)
        ok = not website_text.startswith("[Website")
        status.update(
            label="✅ Website scraped" if ok else "⚠️ Could not scrape website (proceeding with questionnaire data)",
            state="complete" if ok else "error"
        )

    context = build_context(company_name, company_short, info, website_text)

    # ── Step 3: AI content generation ────────────────────────
    ai = {}
    with st.status("🤖 Generating AI content for slides...", expanded=True) as status:

        # Slide 4 – description
        st.write("Slide 4 – Company description paragraph...")
        try:
            ai["slide4_desc"] = groq_call(
                client, SLIDE4_DESC_PROMPT.format(context=context))
        except Exception as e:
            ai["slide4_desc"] = f"{company_name} is a leading specialty glass and pharmaceutical packaging manufacturer..."
            st.warning(f"Slide 4 desc: {e}")

        # Slide 4 – scope paragraph
        st.write("Slide 4 – Scope understanding paragraph...")
        try:
            ai["slide4_scope"] = groq_call(
                client, SLIDE4_SCOPE_PROMPT.format(context=context))
        except Exception as e:
            ai["slide4_scope"] = f"{company_short} seeks support to establish a robust, end-to-end data privacy program aligned with DPDPA 2023."
            st.warning(f"Slide 4 scope: {e}")

        # Slide 4 – scope bullets
        st.write("Slide 4 – Scope bullet points...")
        orig_bullets = "\n".join([
            "Conduct an enterprise‑wide applicability assessment and privacy gap analysis, covering data discovery, lifecycle mapping, inventories, RoPA and documentation of internal/external data flows across EIIL's manufacturing, R&D, supply chain, procurement, commercial, HR, enterprise systems and distribution operations.",
            "Assess privacy, information security and regulatory risks across EIIL's manufacturing, quality, logistics, commercial, enterprise and SaaS platforms, including analytics environments, physical repositories and third‑party networks such as distributors, retailers, logistics partners and service vendors.",
            "Evaluate governance structures, policies and controls covering lawful purpose, consent (where applicable), retention, erasure, grievance handling, DPR workflows, cross‑border transfers and personal data breach processes.",
            "Design and operationalize a scalable privacy governance and risk framework, defining roles, accountability, escalation paths and procedures for DPIAs and risk‑based reviews of new systems, digital initiatives and operational programs.",
            "Support rollout of updated privacy policies, notices and procedures for consent, DPR, retention/deletion, breach response and DPIA processes, tailored for corporate, manufacturing, R&D, commercial and customer‑facing teams.",
            "Coordinate remediation across key platforms to strengthen consent workflows, DPR handling, third‑party data sharing controls, data minimization and privacy‑by‑design requirements with support from selected tooling partners.",
            "Deliver role‑based privacy training, define governance KPIs and RACI structures and enable reporting and dashboards to support continuous oversight, audit readiness, regulatory preparedness and executive visibility.",
        ])
        try:
            raw_bullets = groq_call(
                client,
                SLIDE4_BULLETS_PROMPT.format(context=context, bullets=orig_bullets),
                max_tokens=1800,
            )
            ai["slide4_bullets"] = [l.strip() for l in raw_bullets.split("\n") if l.strip()]
        except Exception as e:
            ai["slide4_bullets"] = []
            st.warning(f"Slide 4 bullets: {e}")

        # Slide 11
        st.write("Slide 11 – Privacy Operating Model paragraph...")
        try:
            ai["slide11_para"] = groq_call(
                client, SLIDE11_PROMPT.format(context=context))
        except Exception as e:
            ai["slide11_para"] = f"For this engagement, the privacy compliance model will be applied exclusively to the internal functions, processes and governance structures of {company_name}, supporting its {', '.join(info['core_business_lines'][:2])} operations and corporate functions."
            st.warning(f"Slide 11: {e}")

        # Slide 17
        st.write("Slide 17 – Data Lifecycle (6 sections)...")
        try:
            raw17 = groq_call(
                client, SLIDE17_PROMPT.format(context=context), max_tokens=2500)
            raw17 = re.sub(r"^```(?:json)?", "", raw17).strip()
            raw17 = re.sub(r"```$", "", raw17).strip()
            ai["slide17_lifecycle"] = json.loads(raw17)
        except Exception as e:
            ai["slide17_lifecycle"] = {}
            st.warning(f"Slide 17: {e}")

        status.update(label="✅ All AI content generated", state="complete")

    # ── Step 4: Modify PPTX ───────────────────────────────────
    with st.status("📝 Applying changes to PPTX...", expanded=False) as status:
        try:
            output_bytes = build_presentation(pptx_bytes, company_name, company_short, info, ai)
            status.update(label="✅ PPTX ready!", state="complete")
        except Exception as e:
            status.update(label="❌ PPTX modification failed", state="error")
            st.error(f"Error: {e}")
            import traceback; st.code(traceback.format_exc())
            st.stop()

    st.success("🎉 Proposal generated successfully!")

    # Preview AI content
    with st.expander("🔍 Preview AI-generated content"):
        with col2:
            st.subheader("🤖 AI-Generated Paragraphs")
        st.markdown("**Slide 4 – Company Description:**")
        st.info(ai.get("slide4_desc",""))
        st.markdown("**Slide 4 – Scope Paragraph:**")
        st.info(ai.get("slide4_scope",""))
        if ai.get("slide4_bullets"):
            st.markdown("**Slide 4 – Scope Bullets:**")
            for b in ai["slide4_bullets"]:
                st.write(f"• {b}")
        st.markdown("**Slide 11 – Operating Model Paragraph:**")
        st.info(ai.get("slide11_para",""))
        if ai.get("slide17_lifecycle"):
            st.markdown("**Slide 17 – Data Lifecycle:**")
            for k, v in ai["slide17_lifecycle"].items():
                st.markdown(f"*{k.replace('_',' ').title()}:* {v}")

    # Download
    safe = re.sub(r"[^\w\s-]","", company_name).strip().replace(" ","_")[:40]
    filename = f"Proposal_Data_Privacy_{safe}_March2026.pptx"
    st.download_button(
        label="⬇️ Download Personalised Proposal",
        data=output_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary",
    )
    st.caption(
        "ℹ️ Sections where the questionnaire had no data, or where AI could not confidently "
        "generate content, are left unchanged from the template for human review."
    )

else:
    st.info("👈 Fill in all details in the sidebar and upload both files, then click **Generate Proposal**.")
    with st.expander("📖 How it works"):
        st.markdown("""
**Two inputs drive the output:**

1. **Template PPTX** — the master proposal file (Eveready template)
2. **Pre-Scoping Questionnaire (.docx)** — filled in for the new company

**What changes per slide:**

| Slide | What changes | Source |
|---|---|---|
| 1 | Company name | Manual input |
| 4 | Company description + scope paragraphs + bullets | AI (Groq) |
| 5 | Employee count, hosting, apps, departments, data subjects | Questionnaire |
| 11 | Privacy Operating Model paragraph | AI (Groq) |
| 12 | EIIL name references in bullets | Auto-replace |
| 14 | EIIL name references in Phase I | Auto-replace |
| 17 | All 6 Data Lifecycle sections | AI (Groq) |
| 19 | EIIL name references in Phase II | Auto-replace |

**Everything else** — all visuals, colours, charts, team bios, methodology diagrams, fees, Protiviti info — stays exactly as in the template.
        """)
