# MERL AI Agent

**Professional AI-Powered Dashboard for Monitoring, Evaluation, Research & Learning**

A clean, modern, fully-featured Streamlit web application that turns your project data into actionable insights — with beautiful charts, an intelligent AI analyst, and one-click professional reports.

![MERL AI Agent](https://via.placeholder.com/1200x600/0D6EFD/FFFFFF?text=MERL+AI+Agent+Dashboard)

---

## ✨ Key Features

- **Beautiful Modern UI** — Professional blue/green design, responsive, easy navigation with a top bar
- **Drag & Drop Upload** — Excel or CSV with flexible column mapping
- **Manual Data Entry** — Simple form to add activities on the fly
- **Powerful Interactive Dashboard** — 6 KPI cards + 4 high-quality Plotly charts (Target vs Actual, Distribution, Status, Regional)
- **Smart Offline AI Assistant** — Ask questions in plain English. Gets real insights, risks, and recommendations from *your* data (no API key needed)
- **One-Click Professional Reports** — Export polished PDF and Word documents with executive summary + recommendations
- **Smart Filters** — Filter by Region and Status across the whole app

---

## 🚀 Run Locally (Windows / macOS / Linux)

```bash
# 1. Go to the project folder
cd MERL-Agent

# 2. (Recommended) Create virtual environment
python -m venv .venv
source .venv/bin/activate          # macOS/Linux
# .\.venv\Scripts\Activate.ps1     # Windows PowerShell

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

The app opens automatically in your browser at `http://localhost:8501`.

---

## ☁️ Deploy to Streamlit Community Cloud (Free Public Link)

### Step-by-step (takes ~5 minutes)

1. **Push the code to GitHub**
   - Create a new **public** GitHub repository
   - Upload / push these files:
     - `app.py`
     - `requirements.txt`
     - `sample_data.csv`
     - `.streamlit/config.toml`
     - `README.md`

2. **Go to Streamlit Cloud**
   - Visit [https://share.streamlit.io](https://share.streamlit.io)
   - Sign in with your GitHub account

3. **Deploy**
   - Click **"New app"**
   - Select your repository
   - Set **Main file path** to: `app.py`
   - Click **Deploy**

4. **Done!**
   - Your app gets a free public link like:
     `https://yourusername-merl-ai-agent.streamlit.app`
   - It stays live 24/7 (free tier)

**Pro tips for Cloud:**
- Sample data loads automatically — great for sharing demos
- 100% offline AI (no external calls) → works perfectly on the free plan
- Visitors can click "Load Sample Data" to reset

---

## 📊 Recommended Data Format

Your Excel/CSV should contain these columns (the app intelligently maps common variations):

| Column     | Description                        | Example                     |
|------------|------------------------------------|-----------------------------|
| Activity   | Activity / intervention name       | "Community Training"        |
| Indicator  | What you are measuring             | "% farmers adopting"        |
| Target     | Planned numeric target             | 85                          |
| Actual     | Achieved value                     | 72                          |
| Unit       | Unit of measure                    | %, people, ha, visits       |
| Status     | Current status                     | On Track / Delayed / Behind |
| Region     | Geographic area                    | Northern / All              |
| Comments   | Qualitative notes                  | "Strong women participation"|

---

## 🛠️ Tech Stack

- **Streamlit** (beautiful UI + instant Cloud hosting)
- **Pandas** + **Plotly** (data processing & professional charts)
- **python-docx** + **fpdf2** (high-quality PDF & Word exports)
- Pure Python intelligent rule-based AI (zero cost, zero API keys)

---

## 📄 Version

**v2.1** — Complete modern UI/UX redesign + Streamlit Cloud ready

Built with care for MEL/MERL professionals who need fast, trustworthy insights without heavy tools.

---

Questions? Just open the app and ask the built-in AI!