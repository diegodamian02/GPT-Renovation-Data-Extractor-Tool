# 📊 GPT Renovation Data Extractor Tool

## 🔹 Project Overview
This project automates the extraction of **renovation details** from **sales walkthrough transcripts** and updates an **Excel spreadsheet** accordingly. It utilizes **Google Gemini AI** to analyze client conversations and extract structured data into relevant categories such as:

- **General Contracting (GC)**
- **Plumbing**
- **Electrical**
- **Millwork**
- **Scope of Work**
- **Permits**

### ✅ Features:
- **Extracts renovation details from DOCX & PDF transcripts**
- **Accurate budget extraction and conversion**
- **Fuzzy matching to update existing spreadsheet rows**
- **Adds new tasks when no existing match is found**
- **Prevents spreadsheet duplication or unnecessary modifications**
- **Ensures structured and reliable data extraction**

---

## ⚙️ Requirements

Before running the project, ensure you have the following installed:

- **Python 3.8+**
- **Required Python Libraries** (install using pip):
  ```bash
  pip install openpyxl python-docx pypdf2 google-generativeai fuzzywuzzy
