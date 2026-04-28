# Younes Daaiki — Portfolio System

A fully data-driven personal portfolio. Edit your info in Excel → run one command → website updates instantly.

---

## 📁 Folder Structure

```
portfolio/
├── index.html              ← The website (open this in browser)
├── data_manager.py         ← Sync tool: Excel ↔ JSON
├── serve.py                ← Local dev server (run this to preview)
├── data/
│   ├── portfolio.json      ← Live database used by the website
│   └── portfolio_editor.xlsx ← Human-friendly editor (edit this!)
├── assets/                 ← Put your project images here
│   └── photo.jpg           ← Your profile photo (replace this)
└── README.md
```

---

## 🚀 Quick Start

### 1. Add your profile photo
Drop your photo as `assets/photo.jpg`

### 2. Preview the site
```bash
python serve.py
```
Then open http://localhost:8080 in your browser.

### 3. Edit your data
Open `data/portfolio_editor.xlsx` in Excel (or LibreOffice).
Edit any sheet: Profile, Skills, Experience, Education, Certifications, Languages, Projects.

### 4. Sync changes to the website
```bash
python data_manager.py excel_to_json
```
Refresh your browser — changes appear instantly.

---

## 📊 Adding Projects

1. Open `data/portfolio_editor.xlsx`
2. Go to the **Projects** sheet
3. Add a new row at the bottom:
   - **id**: leave blank (auto-generated)
   - **title**: project name
   - **category**: e.g. `Power BI`, `Python`, `Analytics`, `Power Platform`, `ERP`
   - **tags**: comma-separated, e.g. `Power BI, DAX, SAP`
   - **description**: 2–3 sentences about the project
   - **images**: filenames from the `assets/` folder, e.g. `dashboard.png, chart.jpg`
   - **embed**: Power BI embed URL, YouTube embed, or any iframe src
   - **link**: external link to the project (optional)
   - **year**: e.g. `2025`
   - **featured**: `TRUE` or `FALSE`

4. Put your image files in the `assets/` folder
5. Run: `python data_manager.py excel_to_json`

---

## 🖼️ Adding Project Images

- Copy your image files into the `assets/` folder
- In the Excel Projects sheet, in the **images** column, write the filenames separated by commas
- Example: `dashboard_screenshot.png, architecture.png`
- Supported formats: `.jpg`, `.jpeg`, `.png`, `.gif`, `.webp`

---

## 🔗 Embedding Power BI / YouTube

In the **embed** column of the Projects sheet, paste:

**Power BI:**
```
https://app.powerbi.com/reportEmbed?reportId=YOUR_REPORT_ID&...
```

**YouTube:**
```
https://www.youtube.com/embed/VIDEO_ID
```

**Any iframe src** works — Google Data Studio, Tableau Public, etc.

---

## 🔄 Workflow Summary

```
Edit Excel  →  python data_manager.py excel_to_json  →  Refresh browser
```

Or the reverse (if you edited JSON directly):
```
python data_manager.py json_to_excel
```

---

## 🌐 Deploying Online

To host your portfolio online for free:

### GitHub Pages
1. Create a GitHub repo
2. Upload all files
3. Go to Settings → Pages → Deploy from branch `main`
4. Your site will be live at `https://yourusername.github.io/portfolio`

### Netlify (drag & drop)
1. Go to https://netlify.com
2. Drag the entire `portfolio/` folder onto the Netlify dashboard
3. Done — live in 30 seconds

---

## ⚙️ Requirements

- Python 3.7+
- `openpyxl` library: `pip install openpyxl`

---

## 📝 Tips

- The website reads `portfolio.json` directly — it's fast and works offline
- Always run `excel_to_json` after editing the Excel file
- Always run `json_to_excel` if you want to re-open the Excel after editing JSON directly
- Project images are loaded lazily — large images are fine
- The site is fully responsive (works on mobile)
