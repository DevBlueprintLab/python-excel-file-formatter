# 📊 Python Excel File Formatter

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)
![Type](https://img.shields.io/badge/Type-Excel%20Automation-green)

---

## 📖 Overview

A Python automation tool that formats Excel spreadsheets automatically.

The program loads an existing workbook, applies professional formatting, inserts summary formulas, adjusts layout, and saves the formatted workbook as a new file without modifying the original.

---

## ⭐ Why this project matters

Formatting Excel files manually is repetitive and time-consuming.

This tool automates common formatting tasks, making spreadsheets cleaner, easier to read, and ready for reporting with a single execution.

---

## 🧠 Features

- 📂 Opens existing Excel workbooks
- ✅ Validates the file before loading
- 💾 Saves a new formatted workbook
- ➕ Automatically adds a **Total** row
- 🧮 Inserts SUM formulas dynamically
- 💲 Formats monetary columns
- 📏 Adjusts column widths
- 🎨 Styles headers with custom fonts and colors
- 📌 Freezes the header row
- 🔍 Enables Excel Auto Filter
- 📐 Centers worksheet content
- 🧱 Adds borders for improved readability

---

## 🖥️ Demo

### Original Workbook

<p align="center">
<img src="images/before.png" width="700">
</p>

### Terminal Output

<p align="center">
<img src="images/terminal.png" width="700">
</p>

### Formatted Workbook

<p align="center">
<img src="images/after.png" width="700">
</p>

---

## ⚙️ Tech Stack

- Python 3
- openpyxl
- pathlib
- sys

---

## 🚀 How to Run

### 1. Clone the repository

```bash
git clone https://github.com/DevBlueprintLab/python-excel-file-formatter.git
```

### 2. Install openpyxl

```bash
pip install openpyxl
```

### 3. Run the script

```bash
python excel_formatter.py
```

### 4. Enter the path to an Excel workbook

Example

```
C:\Users\YourName\Downloads\sales.xlsx
```

The program creates a new file:

```
updated_sales.xlsx
```

---

## 📊 Example Output

```
Excel File Formatter
====================

Please enter the file path:
C:\Users\User\Downloads\sales.xlsx

✓ Workbook formatted successfully!

Saved as:
updated_sales.xlsx
```

---

## 📁 Project Structure

```
python-excel-file-formatter/
├── excel_formatter.py
├── README.md
└── images/
    ├── before.png
    ├── terminal.png
    └── after.png
```

---

## 📚 What I Learned

- Reading existing Excel workbooks
- Working with worksheets dynamically
- Applying fonts, colors and borders
- Formatting cells and numbers
- Creating Excel formulas with Python
- Detecting columns by their headers
- Freezing panes and enabling filters
- Saving modified workbooks safely
- Building reusable Excel automation tools

---

## 🔮 Future Improvements

- Support formatting multiple worksheets
- Detect dates and format them automatically
- Apply conditional formatting
- Automatically resize columns based on content
- Generate charts
- Export formatting logs
- Add a graphical user interface (GUI)

---

Created by **DevBlueprint Lab**
