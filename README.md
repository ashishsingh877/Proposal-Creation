# 📊 AI-Powered Proposal Generator

A Streamlit app that personalises a data-privacy consulting proposal template (`.pptx`) for any new company — using **Groq AI (LLaMA 3.3 70B)** to research the company and replace all company-specific content while keeping the entire template design intact.

---

## ✨ Features

| Feature | Detail |
|---|---|
| Upload any `.pptx` template | Works with the Eveready-style DPDPA proposal template |
| Website scraping | Auto-fetches company info from their homepage |
| Groq AI research | Generates company description, scope, departments, data types etc. |
| Smart PPTX editing | Replaces text **run-by-run**, preserving all fonts, colours & formatting |
| Safe fallback | Sections with insufficient info are left unchanged for human review |
| One-click download | Personalised `.pptx` ready to open in PowerPoint |

---

## 🚀 Quick Start

### 1. Clone the repo
```bash
git clone https://github.com/YOUR_USERNAME/ai-proposal-generator.git
cd ai-proposal-generator
```

### 2. Install dependencies
```bash
pip install -r requirements.txt
```

### 3. Run the app
```bash
streamlit run app.py
```

### 4. Open in browser
Navigate to `http://localhost:8501`

---

## 🔑 Groq API Key

Get a free key at [console.groq.com](https://console.groq.com).  
Enter it in the sidebar when running the app (never stored or logged).

---

## 📁 What Gets Changed

| Slide | What changes |
|---|---|
| **Slide 1** | Company name + title |
| **Slide 4** – Our Understanding | Full company description paragraph |
| **Slide 5** – Scope of Review | Employee count, departments, apps, interfaces |
| **All slides** | Every occurrence of the old company name / abbreviation |

All other slides (methodology, timeline, fees, team, About Protiviti) remain **exactly as in the template**.

---

## 🗂 Project Structure

```
ai-proposal-generator/
├── app.py               # Main Streamlit application
├── requirements.txt     # Python dependencies
└── README.md            # This file
```

---

## ⚙️ Deploying to Streamlit Cloud

1. Push this repo to GitHub  
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**  
3. Select your repo, branch `main`, file `app.py`  
4. Add `GROQ_API_KEY` as a **secret** if you want to pre-fill it  
5. Click **Deploy**

---

## 📝 Notes

- The app does **not** change any images, shapes, colours, or layouts — only text.  
- If the AI cannot confidently fill a field (e.g., exact employee count), it leaves the original template text so a human can update it.  
- Works best when the company has a descriptive homepage.

---

## 🛠 Tech Stack

- [Streamlit](https://streamlit.io) — UI  
- [python-pptx](https://python-pptx.readthedocs.io) — PPTX manipulation  
- [Groq](https://console.groq.com) — LLM API (LLaMA 3.3 70B)  
- [BeautifulSoup4](https://www.crummy.com/software/BeautifulSoup/) — website scraping  
