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
st.title("📊 AI Proposal Generator")
st.markdown(
    "Upload the **template PPTX**, enter the target company details and your Groq key — "
    "the app personalises every relevant slide while keeping design, fonts and layout intact."
)

# ─────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key    = st.text_input("Groq API Key", type="password", placeholder="gsk_...")
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

# ── Slide 4: Full company description paragraph ──────────────
# Must match the original style: 3 rich sentences covering:
# 1) What the company is + product/service segments
# 2) Business model (B2B/B2C/B2B2C) + channels + customer relationships
# 3) Core capabilities + technology + operations + value proposition
P_DESC = """
You are writing a professional consulting proposal for a data-privacy engagement.

Write a single paragraph (3 sentences, 120–140 words total) describing the TARGET COMPANY.
Use the ORIGINAL paragraph below as your structural template — copy the sentence structure exactly,
but replace ALL Eveready/EIIL-specific facts with accurate facts for the new company.

Sentence 1: "[Company] is a [type] [industry] company, operating through [model] spanning [segments/products] across [markets]."
Sentence 2: "The company follows a [B2B/B2C/B2B2C]-driven model, serving [customers/channels] through [distribution model], while maintaining [interface type] through [customer touch points]."
Sentence 3: "[Short name] enables [capabilities] through [technology/systems/processes], ensuring [value proposition]."

TARGET COMPANY:
Name: {company_name}
Short name: {company_short}
Website content: {website_text}

ORIGINAL PARAGRAPH (structure to follow, do NOT copy content):
Eveready Industries India Ltd. (EIIL) is a leading Indian manufacturer of portable energy and \
lighting solutions, operating through a diversified multi-segment model spanning dry-cell batteries, \
flashlights, consumer lighting, professional lighting and electrical accessories across domestic and \
select international markets. The company follows a predominantly B2B and B2B2C-driven model, serving \
distributors, retailers, institutional buyers and large-scale channel partners through one of India's \
widest FMCG-style distribution networks, while maintaining limited B2C interfaces through brand \
engagement, after-sales support and product service programs. EIIL enables end-to-end product \
development, high-volume manufacturing, nationwide distribution and lifecycle management through \
technology-driven quality systems, DSIR-approved R&D capabilities, integrated manufacturing \
facilities and data-enabled supply-chain operations, ensuring safe, reliable, compliant and \
cost-efficient delivery of portable power and lighting solutions across diverse consumer and \
commercial segments.

Return ONLY the rewritten paragraph. Exactly 3 sentences. 120–140 words. No labels or quotes.
"""

# ── Slide 4: Scope paragraph ────────────────────────────────
P_SCOPE = """
Rewrite the sentence below for the target company.
Replace "lending, leasing and factoring operations" with the company's actual business operations.
Keep ALL DPDPA/privacy language word-for-word. Max 50 words.

TARGET COMPANY:
Name: {company_name}, Short: {company_short}
Business context: {website_text}

ORIGINAL:
EIIL seeks support to establish a robust, end-to-end data privacy and personal data protection \
program aligned with the Digital Personal Data Protection Act, 2023 and applicable Rules, \
calibrated to its people, process and technology landscape across lending, leasing and factoring operations.

Return ONLY the rewritten sentence.
"""

# ── Slide 4: 7 scope bullets ────────────────────────────────
P_BULLETS = """
Rewrite the 7 bullet points below for the target company.
Replace EIIL-specific industry terms (manufacturing, R&D, supply chain, distribution, batteries,
flashlights, distributors, retailers, logistics) with equivalent terms for the new company.
Keep ALL privacy/compliance language exactly as-is. Max 35 words per bullet.

TARGET COMPANY:
Name: {company_name}, Short: {company_short}
Business context: {website_text}

BULLETS (return one per line, same order, no numbering, no bullet symbol):
Conduct an enterprise-wide applicability assessment and privacy gap analysis, covering data discovery, lifecycle mapping, inventories, RoPA and documentation of internal/external data flows across EIIL's manufacturing, R&D, supply chain, procurement, commercial, HR, enterprise systems and distribution operations.
Assess privacy, information security and regulatory risks across EIIL's manufacturing, quality, logistics, commercial, enterprise and SaaS platforms, including analytics environments, physical repositories and third-party networks such as distributors, retailers, logistics partners and service vendors.
Evaluate governance structures, policies and controls covering lawful purpose, consent (where applicable), retention, erasure, grievance handling, DPR workflows, cross-border transfers and personal data breach processes.
Design and operationalize a scalable privacy governance and risk framework, defining roles, accountability, escalation paths and procedures for DPIAs and risk-based reviews of new systems, digital initiatives and operational programs.
Support rollout of updated privacy policies, notices and procedures for consent, DPR, retention/deletion, breach response and DPIA processes, tailored for corporate, manufacturing, R&D, commercial and customer-facing teams.
Coordinate remediation across key platforms to strengthen consent workflows, DPR handling, third-party data sharing controls, data minimization and privacy-by-design requirements with support from selected tooling partners.
Deliver role-based privacy training, define governance KPIs and RACI structures and enable reporting and dashboards to support continuous oversight, audit readiness, regulatory preparedness and executive visibility.

Return exactly 7 lines. No numbering, no bullet symbols, no extra text.
"""

# ── Slide 11: Operating model paragraph ─────────────────────
P_S11 = """
You are rewriting one paragraph for a professional consulting proposal slide.

The slide text box has a FIXED size. The paragraph you write MUST be EXACTLY 85 words — \
not 84, not 86. Count carefully before returning.

Rules:
- Replace "Eveready Industries India Ltd." with {company_name}
- Replace "EIIL" with {company_short}
- Replace ALL industry-specific references (manufacturing, supply-chain, distribution, batteries, \
  flashlights, lighting, warranty support, brand engagement, nationwide operational footprint) with \
  accurate equivalents for the new company — its actual business operations, channels, and \
  customer interaction types
- Keep the EXACT sentence structure: Sentence 1 = scope application + operations + channels + \
  B2C data processing touchpoints. Sentence 2 = "This approach ensures..." closing statement.
- Professional consulting prose. Specific to the company. Not generic.

TARGET COMPANY:
Name: {company_name}
Short name: {company_short}
Business context: {website_text}

ORIGINAL (85 words — match this structure and length EXACTLY):
For this engagement, the privacy compliance model will be applied exclusively to the internal \
functions, processes and governance structures of Eveready Industries India Ltd., supporting its \
manufacturing, supply‑chain, commercial, distribution and corporate operations, which primarily \
operate through B2B and B2B2C channels, with limited B2C personal data processing through \
customer service interactions, warranty support and brand engagement programs. This approach \
ensures a focused effort on strengthening EIIL's internal privacy governance and compliance \
capabilities, aligned with applicable regulatory requirements, its operating model and its \
nationwide operational footprint.

Return ONLY the rewritten paragraph. EXACTLY 82 words. No labels, no quotes.
"""

# ── Slide 17: Data Lifecycle ────────────────────────────────
P_S17 = """
You are rewriting 6 Data Lifecycle paragraphs for a professional consulting proposal.
Each paragraph fits inside a FIXED-SIZE text box. You MUST hit the EXACT word count shown.
Count every word carefully. Off by even 2–3 words will cause overflow or empty space.

Rules for every section:
- Replace ALL Eveready/EIIL-specific products, systems, functions, channels and partner types \
  with accurate equivalents for the new company
- Keep the EXACT sentence structure of each original
- Name actual systems, product lines, departments and data types relevant to the new company
- Same formal consulting tone — specific and detailed, never vague or generic

TARGET COMPANY:
Name: {company_name}
Short name: {company_short}
Business context: {website_text}

Return a JSON object with exactly these 6 keys:
"collection", "use_processing", "storage", "sharing", "retention", "disposal"
Value = body paragraph text ONLY (no bold title, no section number prefix).
Return ONLY valid JSON. No markdown fences.

ORIGINALS with EXACT word counts you must match:

collection [EXACTLY 61 words]:
We will review how Eveready Industries India Ltd. (EIIL) collects personal, operational and \
regulatory data across functions such as employee onboarding; manufacturing processes for batteries, \
flashlights and lighting products; distributor onboarding; sales operations; supply‑chain \
coordination; and customer service requirements. This includes data captured through ERP systems, \
plant‑level manufacturing systems, distribution platforms, logistics systems and digital interfaces \
used across EIIL's nationwide network.

use_processing [EXACTLY 64 words]:
We will assess how collected data is used for manufacturing planning, quality control, inventory \
management, supply chain coordination, compliance reporting and performance monitoring across EIIL's \
key segments: batteries, flashlights, consumer lighting, professional lighting and electrical \
accessories. This includes data integration across systems such as ERP, CRM, distributor management \
systems and plant‑level automation platforms, along with tools supporting R&D operations, workforce \
management and operational efficiency.

storage [EXACTLY 48 words]:
We will examine secure storage of manufacturing, safety, employee and vendor data across cloud \
platforms, on‑premise servers at EIIL's manufacturing units, validated production systems, backup \
systems and R&D repositories. Controls for authentication, role‑based access, audit trails and \
compliance with applicable industry and corporate guidelines will also be reviewed.

sharing [EXACTLY 36 words]:
We will evaluate data‑sharing practices with distributors, logistics partners, manufacturing vendors, \
regulatory authorities, retailers and internal teams. This includes reviewing contractual safeguards, \
supply‑chain data‑processing requirements, cross‑border data transfer practices (where applicable), \
anonymization procedures and security measures.

retention [EXACTLY 43 words]:
We will review retention policies for manufacturing logs, quality‑control reports, product testing \
data, R&D records, HR and payroll files, vendor documentation, distributor agreements, operational \
logs and financial documentation. Retention requirements will be assessed against regulatory mandates, \
audit requirements and internal EIIL governance policies.

disposal [EXACTLY 47 words]:
We will verify secure deletion, destruction and anonymization of records across digital platforms, \
manufacturing systems, archival repositories, distributor management systems and physical \
documentation. Disposal workflows will be reviewed for alignment with regulatory expectations and \
internal EIIL data‑governance guidelines to ensure safe and compliant handling of obsolete data.
"""

# ─────────────────────────────────────────────────────────────
# PPTX HELPERS
# ─────────────────────────────────────────────────────────────
def _rep_para(para, rep: dict):
    """Replace text across all runs of a paragraph preserving run[0] formatting."""
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
    TextBox 12 (right column) uses Wingdings § bullets (square appearance).
    TextBox 3  (left column)  uses Arial •  bullets (circle appearance).
    Fix TextBox 12 bullet paragraphs to use the same Arial • style as TextBox 3.
    """
    CIRCLE_BULLET_XML = (
        '<a:buFont xmlns:a="{ns}" typeface="Arial" '
        'panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
        '<a:buChar xmlns:a="{ns}" char="&#8226;"/>'
    ).format(ns=ANS)

    for shape in slide.shapes:
        if shape.name != "TextBox 12":
            continue
        for para in shape.text_frame.paragraphs:
            pPr = para._p.find(f"{{{ANS}}}pPr")
            if pPr is None:
                continue
            # Only fix paragraphs that have a bullet char defined (i.e. actual bullet lines)
            bu_char = pPr.find(f"{{{ANS}}}buChar")
            if bu_char is None:
                continue

            # Remove old buFont and buChar
            for tag in ("buFont", "buChar"):
                el = pPr.find(f"{{{ANS}}}{tag}")
                if el is not None:
                    pPr.remove(el)

            # Insert new Arial circle bullet
            # Parse and append
            new_font = etree.fromstring(
                f'<a:buFont xmlns:a="{ANS}" typeface="Arial" '
                f'panose="020B0604020202020204" pitchFamily="34" charset="0"/>'
            )
            new_char = etree.fromstring(
                f'<a:buChar xmlns:a="{ANS}" char="\u2022"/>'
            )
            pPr.append(new_font)
            pPr.append(new_char)


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
        " EIIL'":   f" {company_short}'",
        " EIIL's":  f" {company_short}'s",
        "(EIIL)":   f"({company_short})",
        "EIIL":     company_short,
    }
    for slide in prs.slides:
        rep_slide(slide, gmap)

    # ── Slide 4 (index 3) ────────────────────────────────────
    if len(prs.slides) > 3:
        s4 = prs.slides[3]

        # Company description paragraph (TextBox 8)
        set_para_text(s4, "TextBox 8", "leading", ai["s4_desc"])

        # Scope paragraph – first para of TextBox 3
        set_para_text(s4, "TextBox 3", "seeks support", ai["s4_scope"])

        # Scope bullets – 7 paragraphs in TextBox 3
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

        # Fix right-side bullets from Wingdings § → Arial •
        fix_slide4_bullets(s4)

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

        st.write("📝 Slide 4 — Company description paragraph…")
        safe("s4_desc",
             P_DESC.format(company_name=company_name,
                           company_short=company_short,
                           website_text=web[:2000]),
             max_tok=350,
             fallback=f"{company_name} is a leading company in its industry, operating across multiple business segments in domestic and international markets. The company follows a B2B-driven model serving enterprise clients and partners through established channels. {company_short} delivers high-quality products and services through integrated operations ensuring safe, compliant and cost-efficient delivery.")

        st.write("📝 Slide 4 — Scope paragraph…")
        safe("s4_scope",
             P_SCOPE.format(company_name=company_name,
                            company_short=company_short,
                            website_text=web[:1500]),
             max_tok=150,
             fallback=f"{company_short} seeks support to establish a robust, end-to-end data privacy and personal data protection program aligned with the Digital Personal Data Protection Act, 2023 and applicable Rules, calibrated to its people, process and technology landscape.")

        st.write("📝 Slide 4 — Scope bullets (7)…")
        try:
            raw_b = groq_call(client,
                              P_BULLETS.format(company_name=company_name,
                                               company_short=company_short,
                                               website_text=web[:1500]),
                              max_tokens=1400)
            lines = [l.strip() for l in raw_b.split("\n") if l.strip()]
            # Remove any accidental numbering / bullet chars the AI added
            lines = [re.sub(r"^[\d]+[.)]\s*", "", l).lstrip("•–-").strip() for l in lines]
            ai["s4_bullets"] = lines[:7]
        except Exception as e:
            ai["s4_bullets"] = []
            st.warning(f"s4_bullets: {e}")

        st.write("📝 Slide 11 — Operating model paragraph…")
        safe("s11",
             P_S11.format(company_name=company_name,
                          company_short=company_short,
                          website_text=web[:2000]),
             max_tok=250,
             fallback=(
                 f"For this engagement, the privacy compliance model will be applied exclusively "
                 f"to the internal functions, processes and governance structures of {company_name}, "
                 f"supporting its core business operations, quality assurance, commercial and corporate "
                 f"functions, which primarily operate through B2B and B2B2C channels, with limited B2C "
                 f"personal data processing through customer service interactions, digital platform usage "
                 f"and product support programs. This approach ensures a focused effort on strengthening "
                 f"{company_short}'s internal privacy governance and compliance capabilities, aligned "
                 f"with applicable regulatory requirements, its operating model and its operational footprint."
             ))
        # Trim s11 to max 82 words to prevent overflow (box fits ~85; 3-word render buffer)
        if ai.get("s11"):
            words = ai["s11"].split()
            if len(words) > 82:
                ai["s11"] = " ".join(words[:82])

        st.write("📝 Slide 17 — Data Lifecycle (6 sections)…")
        # Per-section word limits matching exact box sizes
        S17_LIMITS = {
            "collection": 61, "use_processing": 64, "storage": 48,
            "sharing": 36, "retention": 43, "disposal": 47,
        }
        try:
            raw17 = groq_call(client,
                              P_S17.format(company_name=company_name,
                                           company_short=company_short,
                                           website_text=web[:2000]),
                              max_tokens=3000)
            raw17 = re.sub(r"^```(?:json)?", "", raw17).strip()
            raw17 = re.sub(r"```$", "", raw17).strip()
            ai["s17_lifecycle"] = json.loads(raw17)
            # Trim each section to its exact box word limit
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
        st.markdown("**Slide 4 – Company Description:**")
        st.info(ai.get("s4_desc", ""))
        st.markdown("**Slide 4 – Scope Paragraph:**")
        st.info(ai.get("s4_scope", ""))
        if ai.get("s4_bullets"):
            st.markdown("**Slide 4 – Scope Bullets:**")
            for b in ai["s4_bullets"]:
                st.write(f"• {b}")
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
