# **Parser Launcher — Automatic CEMC Data Processing Tool**

A powerful tool designed to automate and speed up SEMS-related data processing in Uzbekistan.
This launcher significantly reduces the time required to handle large TXT/Excel datasets —
from **up to 3 months of manual work** to **about 10 minutes**, thanks to full automation.

---

## 🚀 Overview

This project processes incoming and outgoing **RRL** and **SPS** datasets.
It automatically reads TXT data (T11, T12, T13 files), extracts the required fields,
and exports processed results into Excel files.

The goal of this project is to:

* automate SEMS data preparation
* reduce processing time from months to minutes
* avoid manual errors
* provide a stable, repeatable, and reliable workflow

---

## 📂 Project Structure

```
project/
│
├── main.py                     # Entry point that launches all parsers
├── config.py                   # Configuration file (input/output paths)
│
├── rrl_incoming_parser.py      # Processes RRL incoming data
├── rrl_outgoing_parser.py      # Processes RRL outgoing data
│
├── sps_incoming_parser.py      # Processes SPS incoming data
├── sps_outgoing_parser.py      # Processes SPS outgoing data
│
├── build_exe.py                # Script that builds ParserLauncher.exe using PyInstaller
└── requirements.txt            # List of project dependencies
```

---

## ⚙️ How It Works

1. **main.py** loads configuration from `config.py`
2. The script reads TXT data (T11/T12/T13) from input folders
3. Corresponding parsers (RRL/SPS, incoming/outgoing) process and filter data
4. Parsed results are exported to Excel files
5. Processing is completed automatically
6. Results are saved to the output folder

---

## 🛠 Installation

### 1. Clone the repository

```bash
git clone https://github.com/yourname/your-repository.git
cd your-repository
```

### 2. Create and activate virtual environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

---

## 🏗 Building the EXE

To create a standalone `.exe`, use:

```bash
python build_exe.py
```

After the build completes, your executable will appear here:

```
dist/ParserLauncher.exe
```

### ⚠️ Important

After building:

1. Copy `config.py` into the `dist/` folder
2. Edit file paths inside `config.py`
3. Run `ParserLauncher.exe`

---

## ⚙️ Configuration (config.py)

Inside this file you can define:

* input path for TXT files
* output path for Excel files
* logging settings
* additional parameters for data processing

Example:

```python
INPUT_FOLDER = r"C:\data\incoming"
OUTPUT_FOLDER = r"C:\data\processed"
```

---

## 📦 Requirements

Project dependencies (located in requirements.txt):

```txt
pandas>=2.0.0
openpyxl>=3.1.0
xlrd>=2.0.1
pyinstaller>=6.0
```

---

## 🧩 Parsers Description

### **rrl_incoming_parser.py**

Processes incoming RRL data: filtering, cleaning, mapping, exporting to Excel.

### **rrl_outgoing_parser.py**

Processes outgoing RRL data.

### **sps_incoming_parser.py**

Processes incoming SPS data.

### **sps_outgoing_parser.py**

Processes outgoing SPS data.

These modules automate repetitive work that previously took weeks or months.

---

## 👤 Author

**Mirzokir Jalilov**, 2025
Developer & creator of the SEMS data automation tool.

---

## 📄 License (MIT)

```
MIT License

Copyright (c) 2025 Mirzokir Jalilov

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
```

