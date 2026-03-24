# JICR-Automation-Generator
Automates JICR Excel report generation using Python.

## 📌 Project Overview

This project is a Python-based automation tool that generates **JICR (Joint Inspection Completion Report)** Excel files in seconds.

It eliminates manual work (which usually takes ~25 minutes per report) and automates the entire process using structured dataset input.

---

## ⚙️ Features

* 📊 Reads data from Excel dataset
* 🧠 Filters data based on **Panchayat & Ward**
* ⚡ Generates multiple JICR reports instantly
* 📄 Automatically fills:

  * FORM 10
  * LUMINARY
  * BATTERY
  * SOLAR
  * DETAILS
* 🎯 Maintains formatting, styles, and structure from template
* 📁 Saves output in organized folder

---

## 🛠️ Technologies Used

* Python 🐍
* pandas
* openpyxl

---

## 📂 Project Structure

```
📁 Project Folder
 ├── DATASET.xlsx
 ├── TEMPLATE.xlsx
 ├── script.py
 ├── JICR_Reports/ (Generated files)
```

---

## ▶️ How to Use

### Step 1: Install dependencies

```bash
pip install pandas openpyxl
```

### Step 2: Place required files

* Add `DATASET.xlsx`
* Add `TEMPLATE.xlsx`

### Step 3: Run the script

```bash
python script.py
```

### Step 4: Provide input

* Enter number of JICR reports
* Enter Panchayat name
* Enter Ward numbers (comma-separated)

---

## 📥 Output

* JICR files will be generated in:

```
JICR_Reports/
```

---

## 🧠 How It Works

* Reads dataset using pandas
* Filters rows based on user input
* Copies template structure using openpyxl
* Inserts filtered data into multiple sheets
* Preserves formatting and footer sections

---

## 🚀 Performance

* Generates report in seconds ⚡
* Can handle large datasets efficiently
* Supports multiple ward processing

---

## 👨‍💻 Author

**Akshay Kumar Sharma**
B.Sc. Computer Science | IIT Patna

---

## 📌 Future Improvements

* GUI interface (no terminal input)
* Auto file selection
* Error logging system
* Batch processing without manual input

---

## ⭐ Note

This project is built to solve real-world manual reporting problems and improve productivity significantly.

---

## 💡 If you like this project

Give it a ⭐ on GitHub!
