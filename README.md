# 📊 Python Excel File Formatter

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)
![Type](https://img.shields.io/badge/Type-Excel%20Automation-green)

---

##  Overview

A Python automation tool that automatically formats Excel spreadsheets into clean, structured, and report-ready files.

It loads an existing workbook, applies professional styling, inserts calculated fields, and exports a formatted version without modifying the original file.

---

## ⭐ Why this project matters

Manually formatting Excel files is repetitive, error-prone, and time-consuming—especially in reporting workflows.

This tool automates formatting tasks and transforms raw spreadsheets into clean, structured, and presentation-ready reports in seconds.

---

## 🧠 Key Features

- 📂 Loads existing Excel workbooks safely  
- ✅ Validates file input before processing  
- 💾 Saves formatted output as a new file  
- ➕ Automatically generates summary (Total) rows  
- 🧮 Inserts dynamic Excel formulas (SUM)  
- 💲 Formats currency and numeric columns  
- 📏 Auto-adjusts column widths  
- 🎨 Applies professional header styling  
- 📌 Freezes header row for usability  
- 🔍 Enables filtering for data analysis  
- 📐 Aligns and structures table layout  
- 🧱 Adds borders for readability  

---

## 🖥️ Demo

### 📄 Original Workbook

<p align="center">
  <img src="images/before.png" width="650"/>
</p>

### ⚙️ Execution Process

<p align="center">
  <img src="images/terminal.png" width="650"/>
</p>

### ✨ Formatted Output

<p align="center">
  <img src="images/after.png" width="650"/>
</p>

---

## 🛠️ Tech Stack

- Python 3  
- openpyxl  
- pathlib  
- sys  

---

## ⚡ How to Run

```bash
git clone https://github.com/DevBlueprintLab/python-excel-file-formatter.git
cd python-excel-file-formatter
pip install openpyxl
python excel_formatter.py
