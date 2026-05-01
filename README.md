# RName - Student Assignment Batch Renaming Tool

**English** | [简体中文](README_zh.md)

<div align="center">

**A Python-based tool for batch renaming student assignments and tracking submissions**

[![Python Version](https://img.shields.io/badge/python-3.11+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey.svg)](https://www.microsoft.com/windows)

</div>

---

## 📋 Table of Contents

- [Introduction](#-introduction)
- [Demo](#️-demo)
- [Key Features](#-key-features)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Usage](#-usage)
- [Renaming Rules](#-renaming-rules)
- [Excel List Requirements](#-excel-list-requirements)
- [Important Notes](#-important-notes)
- [Changelog](#-changelog)
- [FAQ](#-faq)

---

## 📖 Introduction

RName is a student assignment batch renaming tool designed for teachers. It automatically standardizes assignment filenames based on an Excel roster and tracks submission status.

**Author**: Leo (geneleoc@gmail.com)  
**Version**: v1.3.12  
**Updated**: 2023.03.24  
**Created**: 2021.12.28

---

## 🖼️ Demo

<div align="center">
  <img src="assets/rname_demo1.png" alt="RName Demo Screenshot" width="600">
  <p><i>RName application interface</i></p>
</div>

---

## ✨ Key Features

### Core Features

1. **Batch Renaming** - Standardized bulk renaming for files and folders
   - Customizable naming format (serial number, student ID, name, class, etc.)
   - Format example: `017.2022030432.ZhangJia.22CS1.docx`

2. **Intelligent Matching** - Smart assignment matching based on Excel roster
   - Primary matching via student names (automatic Traditional/Simplified Chinese conversion)
   - Duplicate names distinguished by student ID
   - Auto-recognition of international student markers (`*`) and various suffixes

3. **Duplicate Detection** - Detects duplicate submissions with detailed alerts
   - Prevents overwriting errors
   - Provides detailed duplicate submission list

4. **Statistics** - Automatic submission tracking
   - Generates list of students who haven't submitted
   - Auto-writes submission status to Excel (submitted=1, not submitted=0)
   - Supports multiple assignment tracking (Submission1, Submission2...)

### Special Processing

- ✅ Auto conversion: Traditional to Simplified Chinese (using OpenCC)
- ✅ Recognition and handling of international student markers (`*`)
- ✅ Auto-removal of common suffixes: `(Online)`, `(Makeup)`, `(Retake)`, etc.
- ✅ Auto-numbering for duplicate names: `ZhangSan1`, `ZhangSan2`
- ✅ Support for folders as submissions (no extension)

---

## 🛠 Requirements

### System Requirements

- **OS**: Windows 10/11
- **Python**: 3.11 or higher
- **Microsoft Excel**: Required (dependency for xlwings)

---

## 📦 Installation

1. **Install Python**
   
   Download and install the latest version from [Python Official Website](https://www.python.org/downloads/)

2. **Install Dependencies**
   
   Open Command Prompt (CMD) or PowerShell and run:
   
   ```bash
   pip install xlwings opencc 
   ```

### Dependency Description

- **xlwings**: For reading/writing Excel files
- **opencc**: For Traditional/Simplified Chinese conversion

---

## 🚀 Usage

### GUI Operation

1. **Run the Program**
   
   Double-click `RName.py` or run in command line:
   ```bash
   python RName.py
   ```

2. **Specify Assignment Folder**
   
   Click "Specify Directory" button to select the folder containing student assignments

3. **Specify Excel Roster**
   
   Click "Specify Roster" button to select the Excel file with student information

4. **Configure Renaming Rules** (Optional)
   
   - Check/uncheck required fields (serial number, student ID, name)
   - Set prefixes and suffixes
   - Specify file extension (leave empty to keep original)
   - Class position (1-8, leave empty to exclude class)

5. **Start Renaming**
   
   Click "Start Renaming" button to execute

### Naming Examples

Based on configured settings, generated filenames look like:

- `017.2022030432.ZhangJia.docx` (Serial + Student ID + Name)
- `001.22CS1.LiSi.pdf` (Serial + Class + Name)
- `ZhangSan.International.pptx` (Name only with international marker)

---

## 📐 Renaming Rules

The renaming format consists of the following fields (optional):

| Order | Field | Description | Example |
|-------|-------|-------------|---------|
| 1 | Prefix1 | Custom prefix | `-` |
| 2 | Serial | 3-digit number, range: 001-999 | `017` |
| 3 | Prefix2 | Separator (default: `.`) | `.` |
| 4 | Student ID | Read from Excel | `2022030432` |
| 5 | Prefix3 | Separator (default: `.`) | `.` |
| 6 | Name | Read from Excel (Required) | `ZhangJia` |
| 7 | Suffix | Custom suffix | `.Homework` |
| Extra | Class | Can be inserted before positions 1-8 | `22CS1` |
| 8 | Extension | File extension | `.docx` |

**Note**: 
- Name field is required (always checked)
- Folders submitted as assignments have no extension
- Spaces in filenames are automatically removed

---

## 📊 Excel List Requirements

### Table Format

| Column | Name | Required | Description |
|--------|------|----------|-------------|
| Column A | Student ID | ✅ Required | Must be numeric, no duplicates |
| Column B | Name | ✅ Required | Supports Traditional/Simplified, can have `*` marker |
| Column C | Class | ❌ Optional | Read when needed |
| Column D+ | Submission Stats | - | Auto-written by program |

### Detailed Requirements

1. **Sheet**
   - Uses the first sheet by default
   - Sheet name can be specified in the interface

2. **Rows**
   - Row 1 is the header row
   - Data starts from row 2 to the last row

3. **Columns**
   - Data must start from column 1 (Column A)
   - Can contain 2 columns (Student ID + Name) or 3 columns (Student ID + Name + Class)

4. **Data Writing**
   - Statistics start from column 4 (Column D)
   - If D is occupied, automatically finds the next empty column
   - Column header format: `Submission1`, `Submission2`...
   - Data format: Submitted=`1`, Not Submitted=`0`

### Excel Example

```
| Student ID | Name          | Class    | Submission1 | Submission2 |
|------------|---------------|----------|-------------|-------------|
| 2022030432 | Zhang San     | 22CS1    | 1           | 1           |
| 2022030433 | Li Si*        | 22CS1    | 0           | 1           |
| 2022030434 | Wang Wu(Makeup)| 22CS2   | 1           | 0           |
```

---

## ⚠️ Important Notes

### Key Points

1. **Name Matching Priority**
   
   The program matches based on "name" in filenames (students rarely misspell their names):
   - ✅ Filename must contain the student's correct name
   - ❌ Files without names or with incorrect names cannot be renamed

2. **Duplicate Name Handling**
   
   For duplicate names, the filename must include the correct student ID:
   - When multiple students share the same name (e.g., "Zhang San"), distinction is made via student ID
   - Filename example: `ZhangSan-2022030432-Assignment1.docx`

3. **File Hierarchy Limitation**
   
   - ✅ Only processes first-level files/folders in the specified directory
   - ❌ Does not recurse into subdirectories (second-level)

4. **Special Character Processing**
   
   - ✅ `*` before/after names is recognized as international student marker
   - ✅ Auto-removes suffixes: `(Online)`, `(Makeup)`, `(Retake)`, `(Transfer)`, `(Advanced)`, `(Minor)`, etc.
   - ❌ Ethnic minority marker `△` cannot be automatically recognized and must be removed manually

5. **File Status**
   
   - Ensure files to be renamed are not open (especially Word, Excel, etc.)
   - Ensure Excel roster file is not occupied by other programs

6. **Data Validation**
   
   - Student IDs in Excel must be numeric (convertible to integer)
   - Student IDs cannot be duplicated
   - Student ID and Name columns must have the same number of entries

---

## 📝 Changelog

### v1.3.12 - 2023.03.24

**Environment**: Python 3.11.1, Windows 10/11

- 🐛 Fix: Student ID data type validation, pre-check if student ID is numeric
- 🐛 Fix: Unable to process suffixes like "（Online）"

### v1.3.11 - 2023.01.06

**Environment**: Python 3.11.1, Windows 11

- ✨ Improve: Excel column header changed from "Assignment n" to "Submission n"
- ✨ Improve: Optimized print information display

### v1.3.10 - 2023.01.04

**Environment**: Python 3.11.1, Windows 11

- 🐛 Fix: `dupname` dictionary undefined error due to commenting
- 🐛 Fix: Type conversion issue when finding student ID for duplicate names
- 🐛 Fix: Dialog display issue when icon file doesn't exist

### v1.3.9 - 2022.12.26

**Environment**: Python 3.10.4, Windows 10

- 🐛 Fix: String concatenation error due to non-text format in Excel student ID column
- 🐛 Fix: Issue with student ID having `.0` suffix (float to int to str conversion)
- ✨ Improve: Support for multiple bracket suffixes, e.g., `(Transfer)(Online)`
- ✨ Improve: Uses default icon when icon file is missing, no error
- ✨ Improve: Default serial number changed to 3 digits
- ✨ Improve: Default renaming excludes class

### v1.3.8 - 2022.06.03

### v1.3.7 - 2022.06.02

---

## ❓ FAQ

### Q1: "xlwings module not found" error?

**A**: Ensure dependencies are installed:
```bash
pip install xlwings opencc
```

### Q2: "Excel file already open" error?

**A**: Close all open Excel files (including roster file), then retry.

### Q3: Student ID column read error?

**A**: 
- Ensure student ID column is numeric (convertible to integer)
- Check for duplicate student IDs
- Student ID cannot contain text or special characters

### Q4: Some files not renamed?

**A**: Possible reasons:
- Filename doesn't contain student name
- Duplicate names but student ID in filename is incorrect or missing
- File is occupied by another program

### Q5: How to handle ethnic minority marker `△`?

**A**: The program currently cannot auto-recognize `△` marker. It must be manually removed from the Excel roster.

### Q6: Can it run on macOS or Linux?

**A**: Due to dependency on `xlwings` library (requires Microsoft Excel), currently only Windows is supported.

---

## 🔮 TODO

- [ ] Data pre-check: Validate name, student ID, class format before execution
- [ ] Improve error handling: Prevent data errors from causing filename loss
- [ ] Support more special suffix recognition (e.g., `△` ethnic minority marker)
- [ ] Add logging: Record renaming history
- [ ] Support undo: Recover from incorrect renaming
- [ ] Cross-platform support: Consider alternative Excel libraries for macOS and Linux

---

## 📄 License

This project is licensed under the MIT License. See [LICENSE](LICENSE) file for details.

---

## 👤 Author

**Leo**

- Email: geneleoc@gmail.com
- Created: 2021.12.28

---

## 🙏 Acknowledgments

Thanks to the following open source projects:

- [xlwings](https://www.xlwings.org/) - Python Excel automation library
- [OpenCC](https://github.com/BYVoid/OpenCC) - Chinese Traditional/Simplified conversion tool

---

<div align="center">

**If this tool helps you, feel free to Star ⭐**

</div>
