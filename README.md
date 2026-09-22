# Document Generator

A lightweight desktop application that automates the generation of Word documents and PDFs from a CSV database. 

## Installation

### From source

**1. Clone the repository:**
```
git clone https://github.com/akirasy/document-generator.git
cd document-generator
```

**2. Set up a virtual environment (Recommended):**

```
python -m venv venv
source venv/bin/activate  # On Windows use: venv\Scripts\activate
```

**3. Install dependencies:**

```
pip install -r requirements.txt
```

### From Windows executables

**Just download from the releases page**


## How to Use

### 1. Prepare your `.docx` Template

Create a standard Microsoft Word document. Wherever you want dynamic data to appear, type a placeholder tag. (e.g., `<<NAME>>`, `<<IC_NUMBER>>`, `<<DATE>>`). The script will scan paragraphs and tables to replace these tags.

### 2. Prepare your `.csv` Database

Create a CSV file where the **header row exactly matches your template tags**.

* **Crucial:** The script uses the data in the *first and second column* to name the output files.

**Example `data.csv`:**

```
<<NAME>>,<<IC_NUMBER>>,<<COURSE_NAME>>,<<DATE>>
Ahmad Albab,901201-14-5555,Introduction to Python,22 September 2026
Siti Nurhaliza,880101-10-1234,Advanced Data Science,23 September 2026
```

*(This setup will generate `Ahmad Albab - 901201-14-5555.docx`, `Ahmad Albab - 901201-14-5555.pdf`, etc.)*

### 3. Run the App

```
python main.py    # Or just double-click the Windows executable
```

1. Click **Browse** to select your `.docx` template.
2. Click **Browse** to select your `.csv` data source.
3. Click **Auto-Generate DOCX & PDF**.

The app will create an `output` folder right next to `main.py` and neatly place your generated documents inside.
