# 🏗️ Renovation Scope Extraction Tool 📊

This project automates the extraction of renovation details from sales walkthrough transcripts and updates a structured Excel spreadsheet accordingly.

## 📌 Project Overview
Given:
- **A master spreadsheet** outlining the scope of work for a renovation project.
- **Transcripts** of conversations with potential clients detailing renovation requirements.

The tool:
✅ Uses **Google Gemini AI** to extract renovation details from transcripts (DOCX/PDF).  
✅ Matches extracted data with existing entries in the spreadsheet.  
✅ Updates the **budget, proposed work, lead contact, and other details**.  
✅ If no match is found, **appends the task** at the end of the respective sheet.  
✅ Uses **fuzzy matching** to ensure task descriptions align correctly.  

---

## 🛠️ Installation

### **1️⃣ Clone the repository**
```bash
git clone https://github.com/your-username/renovation-scope-extraction.git
cd renovation-scope-extraction
```

### **2️⃣ Set up a virtual environment (Recommended)**
```bash
python -m venv .venv
source .venv/bin/activate  # On macOS/Linux
# OR
.venv\Scripts\activate   # On Windows
```

### **3️⃣ Install dependencies**
```bash
pip install -r requirements.txt
```

### **4️⃣ Set up Google Gemini API key**
Replace `API_KEY` in `app.py` with your **Google Gemini API Key**:
```python
API_KEY = "your-api-key-here"
```

---

## 🚀 How to Run the Project

### **Basic Usage**
To extract renovation details from a transcript and update an existing spreadsheet, run:

```bash
python app.py <transcript_file> <spreadsheet_file>
```

### **Example:**
```bash
python app.py "6030 Collins Ave Summary and Transcript 2025-03-03.docx" "renovation_data.xlsx"
```

### **For Different File Formats**
- **Running with a DOCX file**:
  ```bash
  python app.py "example_transcript.docx" "example_spreadsheet.xlsx"
  ```
- **Running with a PDF file**:
  ```bash
  python app.py "example_transcript.pdf" "example_spreadsheet.xlsx"
  ```

### **What Happens When You Run the Script?**
✅ The script **extracts renovation details** from the provided **DOCX or PDF transcript**.  
✅ It **matches the extracted data** with **existing entries** in the spreadsheet.  
✅ **If a match is found**, it **updates** the budget, proposed work, and lead contact.  
✅ **If no match is found**, it **adds the task** to the end of the respective sheet in the spreadsheet.  
✅ **No new spreadsheets are created**—all updates are saved directly to the provided spreadsheet.

---

## 📌 Notes
⚠️ **Known Issue**: There is a bug affecting **budget updates**, sometimes preventing the correct assignment of values.  
🔹 **Future Enhancements**: Improve budget extraction logic and optimize fuzzy matching for better accuracy.  

---

## 📝 License
COPYRIGHT: Diego Damian 2025
---

## 📩 Contact
For any questions or contributions, feel free to create an issue or pull request on GitHub.