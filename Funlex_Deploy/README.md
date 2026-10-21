# Funlex Cakes and Cream – Sales & Stock Management

**The Taste Is The Difference**

Secure internal web app for recording product sales, tracking stock, and exporting data to Excel.

![Funlex Logo](funlex_logo.png)

---

## Features

| Feature | Description |
|---------|-------------|
| **Access Code Login** | Only authorised staff can open the app |
| **Unique ID** | Every sale & stock movement gets a unique ID |
| **Sales Entry** | Date, Product, Units → auto Amount, Opening & Closing Stock |
| **50+ Products** | Pre-loaded with prices and stock levels |
| **Daily Totals** | Automatic revenue & units summary per day |
| **Stock Tracker** | Production, Purchases, Spoilage, Adjustments |
| **Excel Export** | Download Sales, Movements, Products & Daily Totals |
| **Product Management** | Add products or update prices |

---

## Quick Start (Local)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the app
streamlit run app.py
```

Open the URL shown in the terminal (usually `http://localhost:8501`).

**Default access code:** `FUNLEX2026`  
Change it immediately in **Settings** after first login.

---

## Deploy on Streamlit Community Cloud (Free)

1. Create a **GitHub** account (if you don’t have one).
2. Create a new repository, e.g. `funlex-sales-app`.
3. Upload **all files** in this folder to the repository:
   - `app.py`
   - `requirements.txt`
   - `funlex_logo.png`
   - `.streamlit/config.toml`
   - `README.md`
4. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
5. Click **New app** → select your repository → main file `app.py` → Deploy.
6. Your live app URL will look like:  
   `https://your-username-funlex-sales-app.streamlit.app`

**Important notes for cloud deployment:**
- The SQLite database is created automatically on first run.
- Data is stored on Streamlit’s server (persists between restarts).
- For production use with multiple staff, change the access code and keep the URL private.

---

## Project Structure

```
funlex_app/
├── app.py                 # Main application
├── requirements.txt       # Python dependencies
├── funlex_logo.png        # Official Funlex logo
├── .streamlit/
│   └── config.toml        # Brand theme (pink)
└── README.md
```

---

## Access Code

- Default: `FUNLEX2026`
- Change it under **Settings** inside the app.
- Share the new code only with authorised staff.

---

## Excel Export Sheets

| Sheet | Content |
|-------|---------|
| Sales | Unique ID, Date, Product, Opening, Units, Amount, Closing |
| Stock Movements | Production / Purchase / Spoilage history |
| Products | Current catalogue with prices & stock |
| Daily Totals | Revenue and units per day |

---

© 2026 Funlex Cakes and Cream
