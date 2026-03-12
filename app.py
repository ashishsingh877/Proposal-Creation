"""
AI-Powered Proposal Generator — Protiviti DPDPA Privacy Proposal

Changes per run:
  Slide 1  – Company name
  Slide 4  – AI rewrites company description (detailed 3-sentence paragraph) +
             scope paragraph + 7 scope bullets (word-limited to prevent overflow)
             RIGHT-side bullets fixed to circle • matching left side
  Slide 11 – AI rewrites operating model paragraph
  Slide 12, 14, 19 – Company name auto-replaced
  Slide 17 – AI rewrites all 6 Data Lifecycle sections
  Slide 5  – NOT touched (human fills manually)
"""

import io, json, re
from copy import deepcopy
from lxml import etree
import streamlit as st
import streamlit.components.v1 as components
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from groq import Groq

# ─────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="AI Proposal Generator", page_icon="📊", layout="wide")

# ── Hide "Manage app" bubble — all themes, all devices ──
st.markdown("""
<style>
/* Manage app bar — primary selector */
[data-testid="stStatusWidget"]              { display: none !important; }
[data-testid="stStatusWidget"] *            { display: none !important; }

/* Toolbar top-right (Fork, GitHub) */
[data-testid="stToolbarActions"]            { display: none !important; }
[data-testid="stMainMenuPopover"]           { display: none !important; }

/* Footer */
footer                                      { display: none !important; }
#MainMenu                                   { display: none !important; }
</style>
""", unsafe_allow_html=True)

st.title("📊 AI Proposal Generator")
st.markdown(
    "Upload the **template PPTX**, enter the target company details — "
    "the app personalises every relevant slide while keeping design, fonts and layout intact."
)

# ─────────────────────────────────────────────────────────────
# GROQ API KEY — read from Streamlit secrets first, fall back to manual input
# In Streamlit Cloud: set GROQ_API_KEY in App Settings → Secrets
# Locally: add to .streamlit/secrets.toml as: GROQ_API_KEY = "gsk_..."
# ─────────────────────────────────────────────────────────────
_secret_key = st.secrets.get("GROQ_API_KEY", "") if hasattr(st, "secrets") else ""

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    if _secret_key:
        groq_api_key = _secret_key
        st.success("🔑 Groq API key loaded from Secrets", icon="✅")
    else:
        groq_api_key = st.text_input(
            "Groq API Key",
            type="password",
            placeholder="gsk_...",
            help="Or set GROQ_API_KEY in Streamlit Secrets to avoid entering it every time.",
        )
    st.markdown("---")
    st.header("🏢 Company Details")
    company_name    = st.text_input("Full Company Name",         placeholder="SGD Pharma India Pvt. Ltd.")
    company_short   = st.text_input("Abbreviation / Short Name", placeholder="SGD")
    company_website = st.text_input("Website URL",               placeholder="https://www.sgd-pharma.com")
    st.markdown("---")
    st.header("📁 Template")
    uploaded_ppt = st.file_uploader("Upload Template PPTX", type=["pptx"])
    st.markdown("---")
    generate_btn = st.button("🚀 Generate Proposal", use_container_width=True, type="primary")

# ─────────────────────────────────────────────────────────────
# NAMESPACE SHORTHAND
# ─────────────────────────────────────────────────────────────
ANS = "http://schemas.openxmlformats.org/drawingml/2006/main"

# ─────────────────────────────────────────────────────────────
# SCRAPE WEBSITE
# ─────────────────────────────────────────────────────────────
def scrape_website(url: str, max_chars: int = 6000) -> str:
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=12)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "lxml")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))[:max_chars]
    except Exception as e:
        return f"[Scrape failed: {e}]"

# ─────────────────────────────────────────────────────────────
# GROQ CALL
# ─────────────────────────────────────────────────────────────
GROQ_MODEL = "llama-3.3-70b-versatile"

def groq_call(client: Groq, prompt: str, max_tokens: int = 1000) -> str:
    r = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.25,
        max_tokens=max_tokens,
    )
    return r.choices[0].message.content.strip()

# ─────────────────────────────────────────────────────────────
# AI PROMPTS
# ─────────────────────────────────────────────────────────────

# ── Pre-step 1: Extract rich company profile ─────────────────
# This is the KEY step — all other prompts consume this profile.
# The more accurate this is, the better every slide will be tailored.
P_COMPANY_PROFILE = """
You are a business analyst. Read the company website content below and extract a structured
profile of the company. Be SPECIFIC — use actual names, not generic placeholders.

Company: {company_name} ({company_short})
Website content: {website_text}

Return a JSON object with EXACTLY these keys:

{{
  "industry": "Primary industry sector (e.g. Business Process Outsourcing, Pharmaceuticals, IT Services, FMCG, Banking)",
  "business_model": "One sentence: B2B/B2C/B2B2C split, who they sell to, how they reach clients",
  "service_lines": "Comma-separated list of 4–6 core service lines or product categories",
  "key_sectors": "Comma-separated list of 3–5 industry verticals or client sectors served",
  "key_functions": "Comma-separated list of 5–7 internal business functions (e.g. Operations, Finance, HR, IT, Compliance, Client Delivery)",
  "key_systems": "Comma-separated list of 4–6 technology systems used (ERP, CRM, cloud platforms, analytics tools, etc.)",
  "data_types": "Comma-separated list of 4–6 main personal data categories processed (e.g. employee data, client data, patient data)",
  "partner_types": "Comma-separated list of 3–5 external party types (e.g. clients, vendors, regulators, sub-contractors)",
  "geographic_footprint": "One phrase describing operational geography (e.g. pan-India with global delivery centres)",
  "headcount_scale": "Approximate scale if mentioned (e.g. 50,000+ employees, mid-size, large enterprise)"
}}

Return ONLY valid JSON. No markdown fences. Be specific — use real terms from the website, not generic ones.
"""

# ── Pre-step 2: Extract business model (kept for backward compat) ─
P_BIZ_MODEL = """
Analyse the company website content below and answer in ONE concise sentence (max 20 words):
What is this company's primary business model and channels?

Focus on: B2B / B2B2C / B2C split, who they sell to, and how they reach customers.

Company: {company_name}
Website content: {website_text}

Return ONLY the single sentence. No labels.
"""

# ── Slide 4: Company description BODY ONLY (TextBox 8 R1 — non-bold) ──────
P_DESC = """
You are writing a professional consulting proposal. Write a company description body for the
TARGET COMPANY below. This is NOT a rewrite of a template — generate fresh, accurate content.

CRITICAL RULES:
1. Start DIRECTLY with "is a leading" — do NOT include the company name
2. EXACTLY 3 sentences — no more, no less
3. EXACTLY 118–124 words total — count carefully
4. Use ONLY the company profile provided — do NOT use manufacturing/product terms unless the
   company actually makes physical products
5. Content per sentence:
   - Sentence 1: "is a leading [industry] [company/provider/firm]..." + actual service lines
     or product segments from the profile + markets/geographies served
   - Sentence 2: Business model + who they serve + specific client sectors + how they reach them
     + any B2C touchpoints
   - Sentence 3: Core delivery capabilities + actual technology systems + value proposition closing

TARGET COMPANY PROFILE:
Name: {company_name}
Short name: {company_short}
Industry: {industry}
Business model: {business_model}
Service lines / Products: {service_lines}
Key sectors served: {key_sectors}
Key systems: {key_systems}
Geographic footprint: {geographic_footprint}

Return ONLY the paragraph body. Start with "is a leading". NO company name, NO labels, NO quotes.
"""

# ── Slide 4: Scope operations phrase (TextBox 3 P0 R3) ───────
P_SCOPE_OPS = """
Complete the sentence below for the target company. Fill in [OPS] ONLY.

TEMPLATE:
"and applicable Rules, calibrated to its people, process and technology landscape across [OPS]."

Rules:
- [OPS] = 5–10 words describing the company's ACTUAL core business operations
- Use the company profile — do NOT use generic or incorrect industry terms
- Examples for a BPO: "business process outsourcing, analytics and digital services operations"
- Examples for a bank: "retail banking, corporate lending and digital financial services operations"
- Examples for IT services: "enterprise IT services, cloud solutions and managed services operations"

TARGET COMPANY PROFILE:
Name: {company_name}, Short: {company_short}
Industry: {industry}
Service lines: {service_lines}
Key sectors: {key_sectors}

Return ONLY the complete sentence starting with "and applicable Rules...". No labels, no quotes.
"""

# ── Slide 4: 7 scope bullets ──────────────────────────────────
P_BULLETS = """
You are writing 7 scope bullets for a DPDPA privacy consulting proposal.
Generate these fresh for the TARGET COMPANY — do NOT copy manufacturing/product terms
unless the company actually operates in that industry.

RULES:
- Each bullet max 35 words
- Replace ALL industry-specific references with accurate terms from the company profile below
- Keep ALL privacy/compliance/governance language exactly as-is
- Be specific: name actual functions, platforms, systems and partner types from the profile
- Bullets 1 and 2 MUST reference the company's actual functions, systems and partner types
- Bullets 3–7 keep the same privacy methodology but replace any operational references

TARGET COMPANY PROFILE:
Name: {company_name}, Short: {company_short}
Industry: {industry}
Service lines: {service_lines}
Key sectors: {key_sectors}
Key functions: {key_functions}
Key systems: {key_systems}
Partner types: {partner_types}

BULLET STRUCTURE (rewrite each one — same privacy methodology, company-accurate operations):
1. Enterprise-wide applicability assessment and privacy gap analysis — mention the company's
   actual functions, systems and data flows (not manufacturing/logistics unless relevant)
2. Privacy, information security and regulatory risks — mention actual platforms, SaaS tools,
   analytics environments and the company's specific third-party network types
3. Governance structures — lawful purpose, consent, retention, erasure, grievance, DPR,
   cross-border transfers, breach processes (keep this largely generic/privacy-focused)
4. Scalable privacy governance and risk framework — roles, DPIAs, digital initiatives
5. Privacy policies, notices, consent, DPR, breach response — tailored to actual teams
6. Remediation across key platforms — consent workflows, DPR, data minimization
7. Privacy training, KPIs, RACI, dashboards — oversight and executive visibility

Return exactly 7 lines. No numbering, no bullet symbols, no extra text.
"""

# ── Slide 4: Right-side "How We Will Help" (TextBox 12) ──────
P_S4_RIGHT = """
You are writing bullet points for the "How We Will Help" section of a DPDPA privacy consulting
proposal for the TARGET COMPANY.

RULES:
- Use the company short name where the original says "EIIL"
- Replace ALL industry-specific operational terms with accurate terms for this company
- Keep ALL DPDPA/privacy/governance methodology language EXACTLY as-is
- Match the EXACT word count in brackets for each bullet — fixed-size text boxes
- Be specific to this company's actual industry — do NOT use manufacturing terms for a BPO/IT firm

TARGET COMPANY PROFILE:
Name: {company_name}
Short name: {company_short}
Industry: {industry}
Business model: {business_model}
Service lines: {service_lines}
Key sectors: {key_sectors}

Return a JSON object with EXACTLY these 6 keys: "b1","b2","b3","b4","b5","b6"
Values = text only. No bullet symbols. Return ONLY valid JSON. No markdown fences.

ORIGINALS — rewrite replacing EIIL/Eveready and any industry-specific terms:

b1 [EXACTLY 28 words]:
Enable {company_short}'s transition to sustained compliance with the Digital Personal Data
Protection Act by translating regulatory requirements into a risk-calibrated privacy and
governance framework aligned to business priorities.

b2 [EXACTLY 22 words]:
Define a risk led, high level compliance roadmap addressing material privacy, data protection
and operational gaps across {company_short}'s personal data processing landscape.

b3 [EXACTLY 16 words]:
Establish prioritized remediation themes, sequencing logic and clear accountability structures
to support effective and scalable compliance.

b4 [EXACTLY 25 words]:
Strengthen privacy governance and control architecture by embedding oversight, decision making
and process rigor in line with privacy by design and privacy by default principles.

b5 [EXACTLY 25 words]:
Consolidate key observations and the compliance roadmap into formal deliverables to support
executive oversight, audit readiness and regulatory preparedness, while enhancing stakeholder
trust and transparency.

b6 [EXACTLY 30 words]:
Deliver a risk prioritized remediation roadmap and support governance enablement through a
Privacy Steering Committee, defined KPIs, RACI structures and PMO aligned reporting to
facilitate coordinated implementation and sustained compliance.
"""

# ── Slide 11: Operating model paragraph ──────────────────────
P_S11 = """
You are writing one paragraph for a professional consulting proposal slide.
The text box is FIXED size — the paragraph MUST be EXACTLY 82 words.

RULES:
- Use the company profile to describe ACTUAL operations — do NOT use manufacturing/logistics
  terms unless the company is actually a manufacturer
- Use the EXACT business model for the channel description
- Keep the EXACT 2-sentence structure shown below
- Count every word — must be exactly 82

TARGET COMPANY PROFILE:
Name: {company_name}
Short name: {company_short}
Industry: {industry}
Business model: {business_model}
Service lines: {service_lines}
Key functions: {key_functions}
Geographic footprint: {geographic_footprint}

SENTENCE STRUCTURE TO FOLLOW:
Sentence 1: "For this engagement, the privacy compliance model will be applied exclusively to
the internal functions, processes and governance structures of [company_name], supporting its
[actual operations from service lines/functions], which primarily operate through [channels
from business model], with [B2C data touchpoints if any]."
Sentence 2: "This approach ensures a focused effort on strengthening [company_short]'s internal
privacy governance and compliance capabilities, aligned with applicable regulatory requirements,
its operating model and its [geographic_footprint] operational footprint."

ORIGINAL for reference (82 words):
For this engagement, the privacy compliance model will be applied exclusively to the internal
functions, processes and governance structures of Eveready Industries India Ltd., supporting its
manufacturing, supply-chain, commercial, distribution and corporate operations, which primarily
operate through B2B and B2B2C channels, with limited B2C personal data processing through
customer service interactions, warranty support and brand engagement programs. This approach
ensures a focused effort on strengthening EIIL's internal privacy governance and compliance
capabilities, aligned with applicable regulatory requirements, its operating model and its
nationwide operational footprint.

Return ONLY the paragraph. EXACTLY 82 words. No labels, no quotes.
"""

# ── Slide 17: Data Lifecycle ──────────────────────────────────
P_S17 = """
You are writing 6 Data Lifecycle paragraphs for a professional consulting proposal.
Each paragraph fits inside a FIXED-SIZE text box — you MUST hit the EXACT word count.
Count every word carefully.

CRITICAL: Use the company profile below to name ACTUAL systems, functions, data types and
partner types for this company. Do NOT copy manufacturing/product/logistics terms unless the
company is actually in that industry.

TARGET COMPANY PROFILE:
Name: {company_name}
Short name: {company_short}
Industry: {industry}
Service lines: {service_lines}
Key sectors: {key_sectors}
Key functions: {key_functions}
Key systems: {key_systems}
Data types: {data_types}
Partner types: {partner_types}
Geographic footprint: {geographic_footprint}

Return a JSON object with exactly these 6 keys:
"collection", "use_processing", "storage", "sharing", "retention", "disposal"
Value = body paragraph text ONLY (no title, no section number).
Return ONLY valid JSON. No markdown fences.

STRUCTURE FOR EACH SECTION (follow sentence structure, use company-accurate content):

collection [EXACTLY 61 words]:
"We will review how [company_name] ([company_short]) collects personal, operational and
regulatory data across functions such as [list 4–5 actual functions/processes using semicolons];
and [last function]. This includes data captured through [list 3–4 actual systems] used across
[company_short]'s [geographic] network."

use_processing [EXACTLY 64 words]:
"We will assess how collected data is used for [3–4 actual use cases] across [company_short]'s
key [segments/service lines]: [list 4–5 actual service lines]. This includes data integration
across systems such as [list 3–4 actual systems], along with tools supporting [2–3 actual
functions]."

storage [EXACTLY 48 words]:
"We will examine secure storage of [3–4 actual data types] across cloud platforms, on-premise
servers at [company_short]'s [actual locations/offices], validated [actual system types], backup
systems and [actual repositories]. Controls for authentication, role-based access, audit trails
and compliance with applicable industry and corporate guidelines will also be reviewed."

sharing [EXACTLY 36 words]:
"We will evaluate data-sharing practices with [list 4–5 actual partner types from profile],
regulatory authorities and internal teams. This includes reviewing contractual safeguards,
[industry-specific] data-processing requirements, cross-border data transfer practices
(where applicable), anonymization procedures and security measures."

retention [EXACTLY 43 words]:
"We will review retention policies for [list 5–6 actual record/data types relevant to this
company], vendor documentation, partner agreements, operational logs and financial documentation.
Retention requirements will be assessed against regulatory mandates, audit requirements and
internal [company_short] governance policies."

disposal [EXACTLY 47 words]:
"We will verify secure deletion, destruction and anonymization of records across [list 3–4
actual platform types for this company], archival repositories, [relevant management systems]
and physical documentation. Disposal workflows will be reviewed for alignment with regulatory
expectations and internal [company_short] data-governance guidelines to ensure safe handling
of obsolete data."
"""

# ─────────────────────────────────────────────────────────────
# PPTX HELPERS
# ─────────────────────────────────────────────────────────────
def _rep_para(para, rep: dict):
    """Replace text in each run individually — preserves bold/italic/colour per run.
    Previously this crushed all runs into run[0] which made everything bold.
    Now each run is updated in-place so formatting is never disturbed."""
    for run in para.runs:
        t = run.text
        for k, v in rep.items():
            if k and v is not None and k in t:
                t = t.replace(k, v)
        if t != run.text:
            run.text = t


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


def clean_apostrophes(prs):
    """
    Fix all apostrophe-space gaps across every text run in the presentation.

    Two root causes:
    1. Template typos: e.g. 'EIIL\u2019 privacy' (curly-apos + space, 's' missing)
       → after name replace: 'SGD Pharma\u2019 privacy' — the space survives
    2. Run-split: run[j] ends with company name (trailing space), run[j+1] starts
       with apostrophe → renders as 'SGD Pharma 's'

    Fix strategy:
    a) Within each run: collapse 'word \u2019s' → 'word\u2019s' and 'word \'s' → 'word\'s'
    b) Across run boundaries: if run[j] ends with a letter/space and run[j+1] starts
       with an apostrophe, strip trailing space from run[j]
    """
    APOS = ("\u2019", "\u2018", "'")   # curly-right, curly-left, straight

    def fix_run_text(text):
        for ap in APOS:
            # Pattern: word-char + space + apostrophe → word-char + apostrophe (no space)
            text = re.sub(r'(\w) ' + re.escape(ap), r'\1' + ap, text)
        return text

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                runs = list(para.runs)
                # (a) Fix within each run
                for run in runs:
                    fixed = fix_run_text(run.text)
                    if fixed != run.text:
                        run.text = fixed
                # (b) Fix across run boundaries
                for j in range(len(runs) - 1):
                    for ap in APOS:
                        if runs[j + 1].text.startswith(ap):
                            # Strip trailing space from run[j]
                            if runs[j].text.endswith(" "):
                                runs[j].text = runs[j].text.rstrip(" ")


def set_para_text(slide, shape_name: str, fragment: str, new_text: str) -> bool:
    """Find paragraph whose combined run-text contains fragment; replace with new_text."""
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
# FIX SLIDE 4 BULLET CONSISTENCY
# ─────────────────────────────────────────────────────────────
def fix_slide4_bullets(slide):
    """
    Ensure TextBox 12 right-column bullet paragraphs render as Arial • circles.

    CRITICAL OOXML ORDER inside <a:pPr>:
      spcBef → spcAft → buClr → buSzPct → buFont → buChar → tabLst → defRPr → extLst
    Appending to the end puts bullets AFTER defRPr → PowerPoint ignores them → no dots.
    We must INSERT before defRPr using the element's index in pPr's child list.
    """
    BULLET_TAGS = ("buNone", "buClrTx", "buClr", "buSzTx", "buSzPct",
                   "buSzClamp", "buFont", "buFontTx", "buChar", "buAutoNum")

    for shape in slide.shapes:
        if shape.name != "TextBox 12":
            continue
        # Snapshot paragraphs to a plain list — avoids lxml proxy identity issues
        paras = list(shape.text_frame.paragraphs)
        for i, para in enumerate(paras):
            if i < 2:                               # P0=heading, P1=intro — no bullet
                continue
            pPr = para._p.find(f"{{{ANS}}}pPr")
            if pPr is None:
                continue

            # ── Remove ALL existing bullet-related elements ──────
            for tag in BULLET_TAGS:
                el = pPr.find(f"{{{ANS}}}{tag}")
                if el is not None:
                    pPr.remove(el)

            # ── Build the 4 elements matching the left column ───
            bu_clr  = etree.fromstring(
                f'<a:buClr xmlns:a="{ANS}"><a:srgbClr val="3C3D3E"/></a:buClr>'
            )
            bu_sz   = etree.fromstring(
                f'<a:buSzPct xmlns:a="{ANS}" val="100000"/>'
            )
            bu_font = etree.fromstring(
                f'<a:buFont xmlns:a="{ANS}" typeface="Arial" '
                f'panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
            )
            bu_char = etree.fromstring(
                f'<a:buChar xmlns:a="{ANS}" char="\u2022"/>'
            )

            # ── INSERT before <a:defRPr> so OOXML order is valid ─
            pPr_children = list(pPr)               # snapshot children list
            defRPr = pPr.find(f"{{{ANS}}}defRPr")
            if defRPr is not None:
                ins_idx = pPr_children.index(defRPr)
                # Insert in reverse order so they land in the correct sequence
                for el in (bu_char, bu_font, bu_sz, bu_clr):
                    pPr.insert(ins_idx, el)
            else:
                for el in (bu_clr, bu_sz, bu_font, bu_char):
                    pPr.append(el)

            # ── Ensure hanging-indent attributes are present ─────
            if not pPr.get("marL"):   pPr.set("marL",   "171450")
            if not pPr.get("indent"): pPr.set("indent", "-171450")


# ─────────────────────────────────────────────────────────────
# MAIN BUILD
# ─────────────────────────────────────────────────────────────
def build_presentation(pptx_bytes: bytes, company_name: str,
                       company_short: str, ai: dict) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))

    # ── Global name replacements on every slide ───────────────
    gmap = {
        "Eveready Industries India Ltd. (EIIL)": company_name,
        "Eveready Industries India Ltd":         company_name.split("(")[0].strip().rstrip(","),
        "Eveready Industries":                   company_name.split("(")[0].strip().rstrip(","),
        # Straight apostrophe variants
        " EIIL'":   f" {company_short}'",
        " EIIL's":  f" {company_short}'s",
        # Curly-apostrophe variants (U+2019) — used throughout the template
        " EIIL\u2019s": f" {company_short}\u2019s",
        " EIIL\u2019":  f" {company_short}\u2019s",   # template typo: missing 's' after apostrophe
        "EIIL\u2019s":  f"{company_short}\u2019s",
        "EIIL\u2019 ":  f"{company_short}\u2019s ",   # template typo in same run (no leading space)
        "EIIL\u2019":   f"{company_short}\u2019s",    # bare curly apos without trailing space
        "(EIIL)":        f"({company_short})",
        "EIIL":          company_short,
    }
    for slide in prs.slides:
        rep_slide(slide, gmap)

    # ── Slide 4 (index 3) ────────────────────────────────────
    if len(prs.slides) > 3:
        s4 = prs.slides[3]

        # ── TextBox 8: Top company description ───────────────
        # Run structure: R0(bold)=company name, R1(normal)=body
        # Global replace already updated R0 (EIIL→company_name).
        # Write AI body to R1 ONLY — never touch R0 (preserves bold on name only).
        if ai.get("s4_desc"):
            for shape in s4.shapes:
                if shape.name != "TextBox 8": continue
                runs = shape.text_frame.paragraphs[0].runs
                if len(runs) >= 2:
                    runs[1].text = ai["s4_desc"]
                    for r in runs[2:]: r.text = ""
                break

        # ── TextBox 3 P0: Scope paragraph ────────────────────
        # Run structure: R0(bold)=company short, R1(normal)=fixed DPDPA text,
        #                R2(bold)=fixed "Digital Personal Data Protection Act, 2023",
        #                R3(normal)=" and applicable Rules...landscape across [OPS]."
        # Global replace already updated R0 (EIIL→company_short).
        # ONLY update R3 with AI operations phrase — R1 and R2 stay exactly as-is.
        if ai.get("s4_ops"):
            for shape in s4.shapes:
                if shape.name != "TextBox 3": continue
                runs = shape.text_frame.paragraphs[0].runs
                if len(runs) >= 4:
                    ops_text = ai["s4_ops"].strip()
                    # Ensure it starts with " and applicable Rules..."
                    if not ops_text.lower().startswith("and "):
                        ops_text = "and applicable Rules, calibrated to its people, process and technology landscape across " + ops_text
                    # Prepend space (R3 starts with a space in original)
                    if not ops_text.startswith(" "):
                        ops_text = " " + ops_text
                    # Ensure ends with period
                    if not ops_text.rstrip().endswith("."):
                        ops_text = ops_text.rstrip() + "."
                    runs[3].text = ops_text
                    for r in runs[4:]: r.text = ""
                break

        # ── TextBox 3 P2–P8: Scope bullets ───────────────────
        bullet_frags = [
            "Conduct an enterprise",
            "Assess privacy, information security",
            "Evaluate governance",
            "Design and operationalize",
            "Support rollout",
            "Coordinate remediation",
            "Deliver role",
        ]
        for frag, new_b in zip(bullet_frags, ai.get("s4_bullets", [])):
            if new_b:
                set_para_text(s4, "TextBox 3", frag, new_b)

        # ── TextBox 12: Fix bullets then write AI text ────────
        # P1 (intro) is NEVER rewritten — it's generic and already perfect:
        # "We help embed a trust-first approach...DPDPA and it's Rules."
        # P0=heading (skip), P1=intro (skip), P2–P7=6 bullets (AI)
        fix_slide4_bullets(s4)

        s4r = ai.get("s4_right", {})
        if s4r:
            for shape in s4.shapes:
                if shape.name != "TextBox 12": continue
                paras = list(shape.text_frame.paragraphs)
                # P2..P7 — 6 bullets only (P0 and P1 are never touched)
                for i, key in enumerate(["b1","b2","b3","b4","b5","b6"], start=2):
                    if s4r.get(key) and len(paras) > i:
                        p = paras[i]
                        if p.runs:
                            p.runs[0].text = s4r[key]
                            for r in p.runs[1:]: r.text = ""
                break

    # ── Slide 11 (index 10) ──────────────────────────────────
    if len(prs.slides) > 10:
        set_para_text(prs.slides[10], "Rectangle 1",
                      "For this engagement", ai["s11"])

    # ── Slide 17 (index 16) ──────────────────────────────────
    if len(prs.slides) > 16:
        s17 = prs.slides[16]
        lc  = ai.get("s17_lifecycle", {})
        for shp_name, key in [
            ("Rectangle 41", "collection"),
            ("Rectangle 8",  "use_processing"),
            ("Rectangle 9",  "storage"),
            ("Rectangle 10", "sharing"),
            ("Rectangle 11", "retention"),
            ("Rectangle 12", "disposal"),
        ]:
            if lc.get(key):
                set_para_text(s17, shp_name, "We will", lc[key])

    # ── Slides 12, 14, 19 – global replacements already done ─
    # (gmap handles all remaining EIIL references)

    # ── Final pass: remove all apostrophe-space gaps in every slide ─
    clean_apostrophes(prs)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────
# GENERATE FLOW
# ─────────────────────────────────────────────────────────────
if generate_btn:
    # ── Validation ───────────────────────────────────────────
    errs = []
    if not groq_api_key:    errs.append("🔑 Groq API Key required.")
    if not company_name:    errs.append("🏢 Full company name required.")
    if not company_short:   errs.append("🏷️ Abbreviation required.")
    if not company_website: errs.append("🌐 Website URL required.")
    if not uploaded_ppt:    errs.append("📁 Template PPTX required.")
    for e in errs:
        st.error(e)
    if errs:
        st.stop()

    pptx_bytes = uploaded_ppt.read()
    client     = Groq(api_key=groq_api_key)

    # ── Step 1: Website scrape ────────────────────────────────
    with st.status("🌐 Scraping company website…", expanded=False) as sts:
        web = scrape_website(company_website)
        ok  = not web.startswith("[Scrape")
        sts.update(
            label="✅ Website scraped" if ok
                  else "⚠️ Could not scrape website — proceeding with company name only",
            state="complete" if ok else "error",
        )

    # ── Step 2: AI content generation ────────────────────────
    ai = {}
    def safe(key, prompt, max_tok=1000, fallback=""):
        try:
            ai[key] = groq_call(client, prompt, max_tok)
        except Exception as e:
            ai[key] = fallback
            st.warning(f"{key}: {e}")

    with st.status("🤖 Generating AI content for slides…", expanded=True) as sts:

        # ── Pre-step 1: Extract full company profile ──────────
        st.write("🔍 Extracting company profile from website…")
        profile = {}
        try:
            raw_profile = groq_call(client,
                                    P_COMPANY_PROFILE.format(
                                        company_name=company_name,
                                        company_short=company_short,
                                        website_text=web[:3000]),
                                    max_tokens=600)
            raw_profile = re.sub(r"^```(?:json)?", "", raw_profile).strip()
            raw_profile = re.sub(r"```$", "", raw_profile).strip()
            profile = json.loads(raw_profile)
        except Exception as e:
            st.warning(f"Profile extraction: {e}")

        # Safe accessors with sensible fallbacks
        industry    = profile.get("industry",    "business services")
        biz_model   = profile.get("business_model",
                                  f"Predominantly B2B, serving enterprise clients through direct channels")
        svc_lines   = profile.get("service_lines",  company_short + "'s core service lines")
        key_sectors = profile.get("key_sectors",    "enterprise clients across multiple sectors")
        key_funcs   = profile.get("key_functions",  "Operations, Finance, HR, IT, Compliance, Client Delivery")
        key_systems = profile.get("key_systems",    "ERP, CRM, cloud platforms and analytics tools")
        data_types  = profile.get("data_types",     "employee data, client data, operational data")
        partner_types = profile.get("partner_types","clients, vendors, sub-contractors, regulators")
        geo_footprint = profile.get("geographic_footprint", "pan-India with global operations")
        ai["biz_model"] = biz_model
        ai["profile"]   = profile

        st.write("📝 Slide 4 — Company description body…")
        safe("s4_desc",
             P_DESC.format(company_name=company_name,
                           company_short=company_short,
                           industry=industry,
                           business_model=biz_model,
                           service_lines=svc_lines,
                           key_sectors=key_sectors,
                           key_systems=key_systems,
                           geographic_footprint=geo_footprint),
             max_tok=500,
             fallback=(
                 f"is a leading {industry} company, operating through a diversified model "
                 f"spanning {svc_lines} across domestic and select international markets. "
                 f"{biz_model.rstrip('.')}. "
                 f"{company_short} enables end-to-end service delivery through technology-driven "
                 f"quality systems and integrated operations, ensuring reliable and compliant "
                 f"delivery across diverse client segments."
             ))
        if ai.get("s4_desc"):
            body = ai["s4_desc"].strip()
            for prefix in [company_name, company_short]:
                if body.lower().startswith(prefix.lower()):
                    body = body[len(prefix):].lstrip(" ,")
            if not body.lower().startswith("is "):
                body = "is " + body
            ai["s4_desc"] = " ".join(body.split()[:124])

        st.write("📝 Slide 4 — Scope operations phrase…")
        safe("s4_ops",
             P_SCOPE_OPS.format(company_name=company_name,
                                company_short=company_short,
                                industry=industry,
                                service_lines=svc_lines,
                                key_sectors=key_sectors),
             max_tok=80,
             fallback=f"and applicable Rules, calibrated to its people, process and technology landscape across {company_short}'s {industry.lower()} operations.")

        st.write("📝 Slide 4 — Scope bullets (7)…")
        try:
            raw_b = groq_call(client,
                              P_BULLETS.format(company_name=company_name,
                                               company_short=company_short,
                                               industry=industry,
                                               service_lines=svc_lines,
                                               key_sectors=key_sectors,
                                               key_functions=key_funcs,
                                               key_systems=key_systems,
                                               partner_types=partner_types),
                              max_tokens=1400)
            lines = [l.strip() for l in raw_b.split("\n") if l.strip()]
            lines = [re.sub(r"^[\d]+[.)]\s*", "", l).lstrip("•–-").strip() for l in lines]
            ai["s4_bullets"] = lines[:7]
        except Exception as e:
            ai["s4_bullets"] = []
            st.warning(f"s4_bullets: {e}")

        st.write("📝 Slide 4 — Right-side 'How We Will Help' bullets…")
        S4R_LIMITS = {"b1":28, "b2":22, "b3":16, "b4":25, "b5":25, "b6":30}
        try:
            raw_r = groq_call(client,
                              P_S4_RIGHT.format(company_name=company_name,
                                                company_short=company_short,
                                                industry=industry,
                                                business_model=biz_model,
                                                service_lines=svc_lines,
                                                key_sectors=key_sectors),
                              max_tokens=800)
            raw_r = re.sub(r"^```(?:json)?", "", raw_r).strip()
            raw_r = re.sub(r"```$",          "", raw_r).strip()
            s4r = json.loads(raw_r)
            for k, lim in S4R_LIMITS.items():
                if k in s4r:
                    s4r[k] = " ".join(s4r[k].split()[:lim])
            ai["s4_right"] = s4r
        except Exception as e:
            ai["s4_right"] = {}
            st.warning(f"s4_right: {e}")

        st.write("📝 Slide 11 — Operating model paragraph…")
        safe("s11",
             P_S11.format(company_name=company_name,
                          company_short=company_short,
                          industry=industry,
                          business_model=biz_model,
                          service_lines=svc_lines,
                          key_functions=key_funcs,
                          geographic_footprint=geo_footprint),
             max_tok=250,
             fallback=(
                 f"For this engagement, the privacy compliance model will be applied exclusively "
                 f"to the internal functions, processes and governance structures of {company_name}, "
                 f"supporting its {svc_lines}, which {biz_model.rstrip('.')}, with limited B2C "
                 f"personal data processing through digital platform and client service interactions. "
                 f"This approach ensures a focused effort on strengthening {company_short}'s internal "
                 f"privacy governance and compliance capabilities, aligned with applicable regulatory "
                 f"requirements, its operating model and its {geo_footprint} operational footprint."
             ))
        if ai.get("s11"):
            words = ai["s11"].split()
            if len(words) > 82:
                ai["s11"] = " ".join(words[:82])

        st.write("📝 Slide 17 — Data Lifecycle (6 sections)…")
        S17_LIMITS = {
            "collection": 61, "use_processing": 64, "storage": 48,
            "sharing": 36, "retention": 43, "disposal": 47,
        }
        try:
            raw17 = groq_call(client,
                              P_S17.format(company_name=company_name,
                                           company_short=company_short,
                                           industry=industry,
                                           service_lines=svc_lines,
                                           key_sectors=key_sectors,
                                           key_functions=key_funcs,
                                           key_systems=key_systems,
                                           data_types=data_types,
                                           partner_types=partner_types,
                                           geographic_footprint=geo_footprint),
                              max_tokens=3000)
            raw17 = re.sub(r"^```(?:json)?", "", raw17).strip()
            raw17 = re.sub(r"```$", "", raw17).strip()
            ai["s17_lifecycle"] = json.loads(raw17)
            for k, limit in S17_LIMITS.items():
                if k in ai["s17_lifecycle"]:
                    words = ai["s17_lifecycle"][k].split()
                    if len(words) > limit:
                        ai["s17_lifecycle"][k] = " ".join(words[:limit])
        except Exception as e:
            ai["s17_lifecycle"] = {}
            st.warning(f"s17_lifecycle: {e}")

        sts.update(label="✅ All AI content generated", state="complete")

    # ── Step 3: Build PPTX ───────────────────────────────────
    with st.status("📝 Applying changes to PPTX…", expanded=False) as sts:
        try:
            output = build_presentation(pptx_bytes, company_name, company_short, ai)
            sts.update(label="✅ PPTX ready!", state="complete")
        except Exception as e:
            sts.update(label="❌ Failed", state="error")
            st.error(str(e))
            import traceback; st.code(traceback.format_exc())
            st.stop()

    st.success("🎉 Proposal generated successfully!")

    # ── Preview ──────────────────────────────────────────────
    with st.expander("🔍 Preview AI-generated content"):
        st.markdown("**Extracted Business Model (used in Slides 4 & 11):**")
        st.success(ai.get("biz_model", ""))
        st.markdown("**Slide 4 – Company Description:**")
        st.info(ai.get("s4_desc", ""))
        st.markdown("**Slide 4 – Scope Paragraph:**")
        if ai.get("profile"):
            st.markdown("**🏢 Extracted Company Profile:**")
            st.json(ai["profile"])
        if ai.get("s4_bullets"):
            st.markdown("**Slide 4 – Scope Bullets (Left):**")
            for b in ai["s4_bullets"]:
                st.write(f"• {b}")
        if ai.get("s4_right"):
            st.markdown("**Slide 4 – How We Will Help (Right):**")
            r = ai["s4_right"]
            for k in ["b1","b2","b3","b4","b5","b6"]:
                if r.get(k): st.write(f"• {r[k]}")
        st.markdown("**Slide 11 – Operating Model Paragraph:**")
        st.info(ai.get("s11", ""))
        if ai.get("s17_lifecycle"):
            st.markdown("**Slide 17 – Data Lifecycle:**")
            for k, v in ai["s17_lifecycle"].items():
                st.write(f"**{k.replace('_', ' ').title()}:** {v}")

    # ── Download ─────────────────────────────────────────────
    safe_name = re.sub(r"[^\w\s-]", "", company_name).strip().replace(" ", "_")[:40]
    filename  = f"Proposal_Data_Privacy_{safe_name}_March2026.pptx"
    st.download_button(
        label="⬇️ Download Personalised Proposal",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary",
    )
    st.caption(
        "ℹ️ Slide 5 (Scope of Review) is intentionally left unchanged — "
        "please fill in company-specific details manually."
    )

# ── Landing state ─────────────────────────────────────────────
else:
    st.info("👈 Fill in all details in the sidebar and upload the template, then click **Generate Proposal**.")
    with st.expander("📖 What changes per slide"):
        st.markdown("""
| Slide | Change | Method |
|---|---|---|
| **1** | Company name in title | Auto-replace |
| **4** | Company description (detailed 3-sentence paragraph) + scope paragraph + 7 bullets | Groq AI |
| **4** | Right-side bullet style fixed to ● circle (matching left side) | Auto-fix |
| **5** | **Not changed** — fill manually | Human |
| **11** | Privacy Operating Model paragraph | Groq AI |
| **12, 14, 19** | All EIIL/Eveready name references | Auto-replace |
| **17** | All 6 Data Lifecycle sections | Groq AI |

All visuals, colours, fonts, charts, team slides and Protiviti content remain untouched.
""")
