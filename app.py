import streamlit as st
import os
import io
import json
import copy
import re
import tempfile
from pptx import Presentation
from pptx.util import Pt
import requests
from bs4 import BeautifulSoup
from groq import Groq

# ─────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────
st.set_page_config(
    page_title="AI Proposal Generator",
    page_icon="📊",
    layout="wide",
)

st.title("📊 AI-Powered Proposal Generator")
st.markdown(
    "Upload your template PPT, enter the target company details, and let AI "
    "personalise the proposal — keeping every design element intact."
)

# ─────────────────────────────────────────
#  SIDEBAR – INPUTS
# ─────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    groq_api_key = st.text_input(
        "Groq API Key", type="password", placeholder="gsk_..."
    )
    st.markdown("---")
    st.header("🏢 Company Details")
    company_name = st.text_input(
        "Company Name", placeholder="e.g. Tata Consultancy Services Limited (TCS)"
    )
    company_short = st.text_input(
        "Short Name / Abbreviation", placeholder="e.g. TCS"
    )
    company_website = st.text_input(
        "Company Website", placeholder="e.g. https://www.tcs.com"
    )
    st.markdown("---")
    st.header("📁 Template PPT")
    uploaded_ppt = st.file_uploader("Upload template .pptx", type=["pptx"])

    generate_btn = st.button("🚀 Generate Proposal", use_container_width=True, type="primary")

# ─────────────────────────────────────────
#  HELPER: scrape website text
# ─────────────────────────────────────────
def scrape_website(url: str, max_chars: int = 8000) -> str:
    """Fetch homepage text from the company website."""
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(url, headers=headers, timeout=12)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "lxml")
        # Remove scripts / styles
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        text = soup.get_text(separator=" ", strip=True)
        text = re.sub(r"\s+", " ", text)
        return text[:max_chars]
    except Exception as e:
        return f"[Could not scrape website: {e}]"

# ─────────────────────────────────────────
#  HELPER: ask Groq for company intel
# ─────────────────────────────────────────
GROQ_MODEL = "llama-3.3-70b-versatile"

RESEARCH_PROMPT = """
You are a business analyst. Based on the company information provided below, 
fill in the following JSON fields for a data-privacy consulting proposal.
If you genuinely don't have enough information to fill a field, 
set the value to null (do NOT invent specific numbers or facts).

Company Name: {company_name}
Company Short Name: {company_short}
Website: {company_website}
Website Content (scraped):
{website_text}

Return ONLY a valid JSON object with exactly these keys:

{{
  "company_full_name": "Full legal name with abbreviation in parentheses",
  "company_short": "Abbreviation / ticker",
  "company_one_liner": "One sentence describing what the company does",
  "company_description_paragraph": "3-4 sentence business overview covering: industry, products/services, business model (B2B/B2C/B2B2C), markets served, and key capabilities. Write in a professional, factual tone suitable for a consulting proposal.",
  "scope_understanding": "2-3 sentences describing what personal data challenges DPDPA compliance will involve for THIS company based on their business model",
  "employee_count": "Approximate employee count as a string e.g. '5,000+' or null",
  "hosting_model": "One sentence on data hosting — e.g. 'On-premise', 'Cloud (AWS/Azure)', 'Hybrid' — or null",
  "application_ecosystem": "Comma-separated core enterprise applications used — e.g. 'SAP ERP, Salesforce CRM, Workday HRIS' — or null",
  "customer_facing_interfaces": "Comma-separated customer-facing digital interfaces — e.g. 'website, mobile app, dealer portal, call center' — or null",
  "departments_in_scope": "Comma-separated list of 6-10 key departments that would be in scope for a data privacy review — e.g. 'HR & People Operations, IT & Cybersecurity, Legal & Compliance, Finance, Sales, Marketing, Operations'",
  "data_subjects": "Comma-separated list of data subject categories relevant to this company — e.g. 'Employees, Customers, Vendors, Dealers'",
  "critical_data_types": "Comma-separated list of key personal data types processed — e.g. 'Employment Data, Financial Data, Customer Contact Data, KYC Documents'",
  "key_business_segments": "Comma-separated list of main business segments or product lines",
  "business_model": "B2B, B2C, B2B2C, or combination — one phrase",
  "industry": "Primary industry sector — one phrase",
  "date_prepared": "March 2026"
}}

Return ONLY the JSON. No markdown fences, no explanation.
"""

def research_company(client: Groq, company_name: str, company_short: str,
                     company_website: str, website_text: str) -> dict:
    """Call Groq to research the company and return structured JSON."""
    prompt = RESEARCH_PROMPT.format(
        company_name=company_name,
        company_short=company_short,
        company_website=company_website,
        website_text=website_text,
    )
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=2048,
    )
    raw = response.choices[0].message.content.strip()
    # Strip markdown fences if present
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    return json.loads(raw)


# ─────────────────────────────────────────
#  HELPER: generate slide-4 understanding text
# ─────────────────────────────────────────
SLIDE4_PROMPT = """
You are writing content for a consulting proposal on data privacy (DPDPA compliance).

Target Company:
- Name: {company_full_name}
- Short name: {company_short}
- Industry: {industry}
- Business model: {business_model}
- Business description: {company_description_paragraph}
- Key segments: {key_business_segments}

Rewrite the following template text, replacing ALL references to 
"Eveready Industries India Ltd. (EIIL)" / "EIIL" with the new company, 
and adjusting ALL company-specific facts (industry, products, operations, 
distribution model, capabilities, etc.) to match the new company.

Keep the same professional tone and sentence structure. 
Do NOT change any privacy/compliance-related content.
Keep the length roughly the same.

ORIGINAL TEXT (slide "Our Understanding on Your Entity/Services" main paragraph):
{original_text}

Return ONLY the rewritten paragraph. No labels, no quotes, no extra text.
"""

SCOPE_PROMPT = """
You are writing content for a consulting proposal on data privacy (DPDPA compliance).

Target Company:
- Name: {company_full_name}
- Short name: {company_short}
- Industry: {industry}
- Departments in scope: {departments_in_scope}
- Employee count: {employee_count}
- Hosting model: {hosting_model}
- Application ecosystem: {application_ecosystem}
- Customer-facing interfaces: {customer_facing_interfaces}
- Data subjects: {data_subjects}
- Key segments: {key_business_segments}

Rewrite the following template scope text, replacing ALL Eveready-specific references 
with details relevant to the new company. Replace department names with the new company's 
departments. Replace application names with the new company's apps. Replace employee 
count with the new value (or leave as "—" if unknown).

Keep the same professional tone and structure.

ORIGINAL SCOPE PARAGRAPH:
{original_text}

Return ONLY the rewritten text. No labels, no quotes.
"""

def generate_slide_text(client: Groq, prompt: str) -> str:
    response = client.chat.completions.create(
        model=GROQ_MODEL,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
        max_tokens=1024,
    )
    return response.choices[0].message.content.strip()


# ─────────────────────────────────────────
#  PPTX UTILITIES
# ─────────────────────────────────────────
def replace_text_in_run(run, old: str, new: str):
    """Replace text in a single run."""
    if old in run.text:
        run.text = run.text.replace(old, new)


def replace_in_paragraph(para, replacements: dict):
    """
    Replace text across the full paragraph while preserving run formatting.
    Strategy: consolidate all runs into the first run, apply replacements,
    then clear remaining runs.
    """
    # Build full paragraph text
    full_text = "".join(r.text for r in para.runs)
    
    # Check if any replacement applies
    needs_change = any(old in full_text for old in replacements)
    if not needs_change:
        return
    
    # Apply all replacements
    new_text = full_text
    for old, new in replacements.items():
        if old and new:
            new_text = new_text.replace(old, new)
    
    if new_text == full_text:
        return
    
    # Put the new text into the first run, clear the rest
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""


def replace_in_text_frame(tf, replacements: dict):
    for para in tf.paragraphs:
        replace_in_paragraph(para, replacements)


def replace_in_shape(shape, replacements: dict):
    if shape.has_text_frame:
        replace_in_text_frame(shape.text_frame, replacements)
    # Tables
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_in_text_frame(cell.text_frame, replacements)
    # Group shapes
    if shape.shape_type == 6:  # MSO_SHAPE_TYPE.GROUP
        for s in shape.shapes:
            replace_in_shape(s, replacements)


def replace_in_slide(slide, replacements: dict):
    for shape in slide.shapes:
        replace_in_shape(shape, replacements)


def replace_paragraph_by_old_text(slide, old_fragment: str, new_full_text: str):
    """
    Find the paragraph whose combined text contains old_fragment and 
    replace its content entirely with new_full_text, preserving the 
    format of the first run.
    """
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            full = "".join(r.text for r in para.runs)
            if old_fragment in full:
                if para.runs:
                    para.runs[0].text = new_full_text
                    for r in para.runs[1:]:
                        r.text = ""
                return True
    return False


def modify_presentation(pptx_bytes: bytes, info: dict,
                        slide4_text: str, scope_text: str) -> bytes:
    """
    Load the PPTX from bytes, apply all company-specific replacements,
    return modified PPTX as bytes.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    company_full = info.get("company_full_name") or ""
    company_short = info.get("company_short") or ""

    # ── Simple text replacements (company name everywhere) ──────────────
    simple_replacements = {
        "Eveready Industries India Ltd. (EIIL)": company_full,
        "Eveready Industries India Ltd": company_full.split("(")[0].strip(),
        "EIIL": company_short,
        "March 2026": info.get("date_prepared", "March 2026"),
    }
    # Also replace short-form standalone if different
    if company_short and company_short != "EIIL":
        simple_replacements["EIIL"] = company_short

    for slide in prs.slides:
        replace_in_slide(slide, simple_replacements)

    # ── Slide 4: company description paragraphs ──────────────────────────
    # The big description paragraph starts with "Eveready Industries India Ltd. (EIIL) is"
    SLIDE4_FRAGMENT = "is a leading Indian manufacturer"
    if len(prs.slides) >= 4:
        slide4 = prs.slides[3]  # 0-indexed
        # After simple replacements it may start differently; search for the fragment
        for shape in slide4.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                full = "".join(r.text for r in para.runs)
                # Match the big description paragraph
                if ("leading Indian manufacturer" in full or
                        "leading" in full and len(full) > 200):
                    if para.runs and slide4_text:
                        para.runs[0].text = slide4_text
                        for r in para.runs[1:]:
                            r.text = ""
                    break

    # ── Slide 5: scope paragraph ──────────────────────────────────────────
    # Update employee count
    emp_count = info.get("employee_count") or "—"
    SLIDE5_FRAGMENT = "1200+"
    if len(prs.slides) >= 5:
        slide5 = prs.slides[4]
        replace_in_slide(slide5, {
            "1200+": emp_count,
            "Audit coverage includes assessment of data handling practices within 8 departments such as HR & People Operations, IT & Cybersecurity, Legal & Compliance , Sales, Distribution, Customer Support, Manufacturing Operations, Marketing": scope_text,
        })
        # Replace application ecosystem text
        app_eco = info.get("application_ecosystem")
        if app_eco:
            replace_in_slide(slide5, {
                "HRIS (Human Resource Information System), CRM (Customer Relationship Management), ERP (Enterprise Resource Planning)": app_eco,
            })
        # Replace customer-facing interfaces
        cf = info.get("customer_facing_interfaces")
        if cf:
            replace_in_slide(slide5, {
                "Web Portal, Dealer/Distributor Portals, Call Center / Customer Support, SAP, Sales force": cf,
            })

    # ── Save and return bytes ─────────────────────────────────────────────
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────
#  ORIGINAL TEMPLATE TEXT (for prompts)
# ─────────────────────────────────────────
ORIGINAL_SLIDE4_PARA = (
    "Eveready Industries India Ltd. (EIIL) is a leading Indian manufacturer of portable energy and "
    "lighting solutions, operating through a diversified multi‑segment model spanning dry‑cell batteries, "
    "flashlights, consumer lighting, professional lighting and electrical accessories across domestic and "
    "select international markets. The company follows a predominantly B2B and B2B2C‑driven model, serving "
    "distributors, retailers, institutional buyers and large‑scale channel partners through one of India's "
    "widest FMCG‑style distribution networks, while maintaining limited B2C interfaces through brand "
    "engagement, after‑sales support and product service programs. EIIL enables end‑to‑end product "
    "development, high‑volume manufacturing, nationwide distribution and lifecycle management through "
    "technology‑driven quality systems, DSIR‑approved R&D capabilities, integrated manufacturing "
    "facilities and data‑enabled supply‑chain operations, ensuring safe, reliable, compliant and "
    "cost‑efficient delivery of portable power and lighting solutions across diverse consumer and "
    "commercial segments."
)

ORIGINAL_SCOPE_PARA = (
    "Audit coverage includes assessment of data handling practices within 8 departments such as "
    "HR & People Operations, IT & Cybersecurity, Legal & Compliance, Sales, Distribution, "
    "Customer Support, Manufacturing Operations, Marketing"
)


# ─────────────────────────────────────────
#  MAIN GENERATE FLOW
# ─────────────────────────────────────────
if generate_btn:
    # Validate inputs
    errors = []
    if not groq_api_key:
        errors.append("🔑 Please enter your Groq API Key.")
    if not company_name:
        errors.append("🏢 Please enter the company name.")
    if not company_short:
        errors.append("🏷️ Please enter the company short name / abbreviation.")
    if not company_website:
        errors.append("🌐 Please enter the company website URL.")
    if not uploaded_ppt:
        errors.append("📁 Please upload the template PPTX file.")

    if errors:
        for e in errors:
            st.error(e)
        st.stop()

    pptx_bytes = uploaded_ppt.read()
    client = Groq(api_key=groq_api_key)

    col1, col2 = st.columns(2)

    with col1:
        with st.status("🔍 Researching company...", expanded=True) as status:
            st.write("Scraping website...")
            website_text = scrape_website(company_website)
            if website_text.startswith("[Could not"):
                st.warning(f"Website scraping failed — proceeding with AI knowledge only.\n\n{website_text}")
                website_text = ""

            st.write("Asking Groq for company intelligence...")
            try:
                info = research_company(
                    client, company_name, company_short, company_website, website_text
                )
                status.update(label="✅ Research complete!", state="complete")
            except Exception as e:
                status.update(label="❌ Research failed", state="error")
                st.error(f"Error during research: {e}")
                st.stop()

    with col2:
        st.subheader("📋 Extracted Company Info")
        display_info = {k: v for k, v in info.items() if v is not None}
        st.json(display_info)

    st.markdown("---")

    with st.status("✍️ Generating personalised slide content...", expanded=True) as status:
        # Generate slide 4 description
        st.write("Generating 'Our Understanding' paragraph (Slide 4)...")
        try:
            slide4_prompt = SLIDE4_PROMPT.format(
                company_full_name=info.get("company_full_name", company_name),
                company_short=info.get("company_short", company_short),
                industry=info.get("industry", ""),
                business_model=info.get("business_model", ""),
                company_description_paragraph=info.get("company_description_paragraph", ""),
                key_business_segments=info.get("key_business_segments", ""),
                original_text=ORIGINAL_SLIDE4_PARA,
            )
            slide4_text = generate_slide_text(client, slide4_prompt)
        except Exception as e:
            st.warning(f"Could not generate slide 4 text: {e}. Will use simple substitution.")
            slide4_text = ORIGINAL_SLIDE4_PARA.replace(
                "Eveready Industries India Ltd. (EIIL)", company_name
            ).replace("EIIL", company_short)

        # Generate slide 5 scope text
        st.write("Generating 'Scope of Review' text (Slide 5)...")
        try:
            scope_prompt = SCOPE_PROMPT.format(
                company_full_name=info.get("company_full_name", company_name),
                company_short=info.get("company_short", company_short),
                industry=info.get("industry", ""),
                departments_in_scope=info.get("departments_in_scope", ""),
                employee_count=info.get("employee_count") or "—",
                hosting_model=info.get("hosting_model") or "—",
                application_ecosystem=info.get("application_ecosystem") or "—",
                customer_facing_interfaces=info.get("customer_facing_interfaces") or "—",
                data_subjects=info.get("data_subjects") or "—",
                key_business_segments=info.get("key_business_segments") or "—",
                original_text=ORIGINAL_SCOPE_PARA,
            )
            scope_text = generate_slide_text(client, scope_prompt)
        except Exception as e:
            st.warning(f"Could not generate scope text: {e}. Will use simple substitution.")
            depts = info.get("departments_in_scope") or "HR & People Operations, IT & Cybersecurity, Legal & Compliance, Finance, Sales, Marketing, Operations"
            scope_text = f"Audit coverage includes assessment of data handling practices across key departments such as {depts}"

        status.update(label="✅ Content generation complete!", state="complete")

    # Preview generated text
    with st.expander("🔍 Preview Generated Slide Content"):
        st.markdown("**Slide 4 – Company Understanding Paragraph:**")
        st.info(slide4_text)
        st.markdown("**Slide 5 – Scope Paragraph:**")
        st.info(scope_text)

    # Modify PPTX
    with st.status("📝 Applying changes to PPTX template...", expanded=False) as status:
        try:
            output_bytes = modify_presentation(pptx_bytes, info, slide4_text, scope_text)
            status.update(label="✅ PPTX ready!", state="complete")
        except Exception as e:
            status.update(label="❌ PPTX modification failed", state="error")
            st.error(f"Error modifying PPTX: {e}")
            st.stop()

    st.success("🎉 Proposal generated successfully!")

    # Clean company name for filename
    safe_name = re.sub(r"[^\w\s-]", "", company_name).strip().replace(" ", "_")[:40]
    filename = f"Proposal_Data_Privacy_{safe_name}_March2026.pptx"

    st.download_button(
        label="⬇️ Download Personalised Proposal",
        data=output_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True,
        type="primary",
    )

    st.markdown("---")
    st.caption(
        "ℹ️ Sections where AI didn't have enough information are left unchanged "
        "from the template — a human reviewer should fill those in."
    )

else:
    # Landing state
    st.info(
        "👈 Fill in the company details in the sidebar and upload the template PPT, "
        "then click **Generate Proposal**."
    )
    with st.expander("📖 How it works"):
        st.markdown("""
1. **Upload** your template `.pptx` proposal file  
2. **Enter** the target company name, abbreviation, and website URL  
3. **Add** your Groq API key  
4. Click **Generate Proposal** — the app will:
   - Scrape the company website for context  
   - Use Groq AI (LLaMA 3.3 70B) to research the company  
   - Replace all company-specific text in the proposal  
   - Keep the entire template design, layout & formatting intact  
5. **Download** the personalised `.pptx` file  

> ⚠️ Where AI doesn't have enough information, the original placeholder text is kept so a human can fill it in manually.
        """)
