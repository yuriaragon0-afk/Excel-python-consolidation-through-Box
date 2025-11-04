[README_ENGLISH.md](https://github.com/user-attachments/files/23345723/README_ENGLISH.md)
# 📊 Excel Report Consolidation Automation (Python + xlwings)

This project automates the **consolidation of multiple Excel reports**, extracting data from specific sheets, merging everything into a single master file, and executing **VBA macros** to generate a formatted final report.

The project was inspired by a real-world automation used in a corporate financial environment, but **all paths and data have been replaced with generic, fictional examples** to protect company confidentiality.

---

## 🚀 Features

- 🔄 Automatic consolidation of multiple Excel files  
- 📑 Reading specific sheets and cell ranges  
- 📦 Merging data into a single master file  
- ⚙️ Executing Excel macros directly from Python  
- 📈 Generating a final consolidated report  

---

## 🧰 Technologies Used

- **Python 3.x**
- pandas — data manipulation  
- openpyxl — reading and writing Excel files  
- xlwings — Excel automation and macro execution  

---

## ⚙️ Project Structure

```
your-project/
├── BPT bridges dummy.ipynb      # Main notebook (automation code)
├── requirements.txt             # Project dependencies
├── .gitignore                   # Git ignored items
└── README.md                    # This file :)
```

---

## 🧩 How to Run the Project

### 1️⃣ Clone the repository

```bash
git clone https://github.com/yuriaragon0-afk/Excel-python-consolidation-through-Box
cd Excel-python-consolidation-through-Box
```

### 2️⃣ Create a virtual environment (optional)

```bash
python -m venv venv
venv\Scripts\activate   # Windows
# or
source venv/bin/activate   # macOS/Linux
```

### 3️⃣ Install dependencies

```bash
pip install -r requirements.txt
```

### 4️⃣ Configure paths in the code

At the beginning of the notebook (or Python script), edit the paths according to your local setup:

```python
folder_path = r"C:/Example/Reports/"
consolidation_path = r"C:/Example/Consolidated/consolidated.xlsx"
source_sheet_name = "Summary"
macro_name = "RunConsolidation"
```

These paths indicate where the input files are located and where the consolidated output will be saved.

---

## ▶️ Execution

If you’re using the **notebook**:
1. Open `BPT bridges dummy.ipynb` in Jupyter or VSCode  
2. Run all cells sequentially  

To convert it into a **Python script**:
```bash
python consolidacao.py
```

The script will:
1. Read all Excel files from the specified folder  
2. Copy data from the defined sheets  
3. Consolidate everything into a single file  
4. Execute the selected macro  
5. Generate the final formatted report  

---

## 📦 Example Data Structure

```
data/
├── report_analyst1.xlsx
├── report_analyst2.xlsx
└── report_analyst3.xlsx
```

---

## 💡 Optional: Using a `.env` File

To make the code more flexible and secure (a professional best practice),  
you can store paths and sheet names in a `.env` file and load them with the `python-dotenv` library:

```python
from dotenv import load_dotenv
import os

load_dotenv()

folder_path = os.getenv("FOLDER_PATH")
consolidation_path = os.getenv("CONSOLIDATION_PATH")
source_sheet_name = os.getenv("SOURCE_SHEET_NAME")
macro_name = os.getenv("MACRO_NAME")
```

Example `.env` file:
```
FOLDER_PATH=C:/Example/Reports/
CONSOLIDATION_PATH=C:/Example/Consolidated/consolidated.xlsx
SOURCE_SHEET_NAME=Summary
MACRO_NAME=RunConsolidation
```

This step is **optional** — the script also works with hardcoded paths.

---

## ⚠️ Confidentiality Notice

This project was inspired by a real corporate automation; however, **all data, names, and paths have been replaced with generic examples**.  
No sensitive or confidential content is included in this repository.

---

## 👤 Author

**[Yuri Aragon]**  
Financial Analyst | Python | Excel | Process Automation  
📧 [yuriaragon0@gmail.com]  
🌐 [https://www.linkedin.com/in/yuriaragon/]

---

## 🏷️ License

Distributed under the MIT License. See the `LICENSE` file (optional) for more information.
