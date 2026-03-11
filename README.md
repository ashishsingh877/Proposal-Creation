# 📊 AI Proposal Generator — Protiviti DPDPA Privacy Proposal

Personalises a data-privacy consulting proposal (`.pptx`) for any new company using:
- A **filled Pre-Scoping Questionnaire** (`.docx`) for factual company data
- **Groq AI (LLaMA 3.3 70B)** for company-specific paragraphs
- Simple find-and-replace for name/date references everywhere else

Design, layout, fonts, colours and all non-company content remain 100% intact.

---

## 🗺️ Slide-by-Slide Changes

| Slide | What changes | How |
|---|---|---|
| **1** | Company name in title | Auto-replace |
| **4** | Company description, scope paragraph, 7 scope bullets | Groq AI |
| **5** | Employee count, hosting model, applications, departments, data subjects | Questionnaire |
| **11** | Privacy Operating Model paragraph | Groq AI |
| **12** | All EIIL name references in bullets | Auto-replace |
| **14** | All EIIL name references in Phase I | Auto-replace |
| **17** | All 6 Data Lifecycle sections (Collection → Disposal) | Groq AI |
| **19** | All EIIL name references in Phase II | Auto-replace |

---

## 🚀 Quick Start

```bash
git clone https://github.com/YOUR_USERNAME/ai-proposal-generator.git
cd ai-proposal-generator
pip install -r requirements.txt
streamlit run app.py
```

Open `http://localhost:8501`

---

## 📝 Pre-Scoping Questionnaire Format

The app parses the standard Protiviti Pre-Scoping Privacy Questionnaire (`.docx`).  
It reads five sections automatically:

| Section | Fields extracted |
|---|---|
| Organisational Overview | Employee count, subsidiaries |
| Governance & Accountability | Policy framework status |
| Business Lines & Stakeholders | Core business lines, departments |
| Data Ecosystem | Applications, interfaces, hosting |
| Data Subjects & Data Types | Data subject categories, data types |

---

## 🔑 Groq API Key

Free key at [console.groq.com](https://console.groq.com).  
Enter in sidebar — never stored.

---

## ☁️ Deploy to Streamlit Cloud

1. Push repo to GitHub  
2. [share.streamlit.io](https://share.streamlit.io) → New app → select `app.py`  
3. Add secret `GROQ_API_KEY` if you want to pre-fill the key  

---

## 🛠 Tech Stack

| Library | Purpose |
|---|---|
| `streamlit` | UI |
| `python-pptx` | PPTX text replacement (run-level, preserves formatting) |
| `python-docx` | Parse questionnaire tables |
| `groq` | LLaMA 3.3 70B for AI paragraphs |
| `beautifulsoup4` | Website scraping for extra context |
