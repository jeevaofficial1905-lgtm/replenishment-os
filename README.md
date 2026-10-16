# ReplenishmentOS & Data Pipeline

A high-performance apparel replenishment and inventory projection engine built for omnichannel retail (EBO, D2C, LFR).

---

## 🚀 Deploy the Entire Application on Render.com

You can host both the **frontend dashboard and backend converter** as a single service on Render for free:

### Method 1: Automatic Blueprint (Recommended)
1. Push this repository to **GitHub**.
2. Go to [Render.com](https://render.com) and click **New +** → **Blueprint**.
3. Select your repository.
4. Render will read `render.yaml` automatically and configure everything:
   - **Environment**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python server.py`
5. Click **Apply**. In 1–2 minutes, your live URL (`https://your-app.onrender.com`) is ready!

### Method 2: Manual Web Service
1. On [Render.com](https://render.com), click **New +** → **Web Service**.
2. Connect your GitHub repository.
3. Configure the following settings:
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `python server.py`
   - **Plan**: `Free`
4. Click **Create Web Service**.

When you visit your Render URL, the complete **ReplenishmentOS dashboard** loads with in-browser data caching and full server-side processing capabilities.

---

## 💻 Standalone Windows Executable (.exe)
Double-click `ReplenishmentOS.exe` to launch the complete native Windows desktop application with zero configuration or dependencies.

---

## 📁 Repository Structure

```text
├── server.py             # Fullstack Render web server (serves frontend + API)
├── render.yaml           # 1-click Render Blueprint specification
├── index.html            # Complete ReplenishmentOS dashboard interface
├── ReplenishmentOS_files/ # Offline vendor dependencies (Chart.js, SheetJS)
├── app.py                # Streamlit alternative converter interface
├── process_pipeline.py   # CLI / automated batch conversion script
├── requirements.txt      # Python dependencies (openpyxl, streamlit)
├── .gitignore            # Git ignore configuration
└── README.md             # Documentation & setup guide
```
