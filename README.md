# ReplenishmentOS & Data Pipeline

A high-performance apparel replenishment and inventory projection engine built for omnichannel retail (EBO, D2C, LFR).

---

## 🚀 1-Click Online Deployment (Streamlit Cloud)

You can run the data converter online for free:

1. Push this repository to **GitHub**.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Select your repository and set `Main file path` to **`app.py`**.
4. Click **Deploy**.
5. You now have a live public or team URL where anyone can drag-and-drop raw Excel files and download `pipeline_data.json` in seconds!

---

## 💻 Local Desktop Usage

### Option A: Local Python / Batch Script
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
2. Double-click `run_pipeline.bat` or run:
   ```bash
   python process_pipeline.py "path/to/raw_pipeline.xlsx"
   ```
3. It will generate `pipeline_data.json` and `PIPELINE WORKING.xlsx`.

### Option B: Local Streamlit App
Run the web converter locally:
```bash
streamlit run app.py
```

---

## 📊 Using the Replenishment Frontend

1. Double-click **`index.html`** (or `ReplenishmentOS.html`) to open it in Chrome, Edge, or Safari.
2. Click **Refresh / Upload** in the top-right corner.
3. Select your generated **`pipeline_data.json`** (or `PIPELINE WORKING.xlsx`).
4. All **45,000+ SKUs** load in **~0.15 seconds** with full stock projection, style health scores, and replenishment recommendations.

---

## 📁 Repository Structure

```text
├── app.py                # Streamlit web app for online cloud conversion
├── index.html            # Standalone ReplenishmentOS frontend (302 KB)
├── process_pipeline.py   # CLI / automated batch conversion script
├── run_pipeline.bat      # Windows 1-click batch runner
├── requirements.txt      # Python dependencies (streamlit, openpyxl)
├── .gitignore            # Git ignore configuration
└── README.md             # Documentation & setup guide
```
