# Legacy Case Note Conversion Tool Chain

## License
**Copyright (C) 2025  Roza Nil**

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <https://www.gnu.org/licenses/>.

---

## Overview

This suite automates the process of extracting, filtering, and structuring client case notes from a master ZIP archive containing legacy Microsoft Word (`.docx`) files. The entire process is managed via the `golive.sh` script for one-command execution.

### Known Limitations (Read Before Running)
**Manual Update Required:** Case managers will need to update the client profile page for current participants manually. Parsing profile data from legacy `.docx` files proved unreliable.
> *These scripts are smart, but not that smart.*

---

## 1. Prerequisites and Setup

### System Requirements
* **Python 3.x**
* **Unix-like environment** (Linux, macOS, or Windows Subsystem for Linux/Git Bash) to run the shell script.

### Library Requirements
The following Python libraries are required:
* `python-docx`
* `pandas`
* `xlsxwriter`
* `openpyxl`

To install these dependencies, run:
```bash
pip install python-docx pandas xlsxwriter openpyxl
```
The 'golive.sh' script automates the entire three-step process:

1.  **Filter:** Scans a master ZIP, keeps only verified client case files (those containing "Entry Date" or "Exit Date" keywords).
2.  **Extract:** Reads the filtered DOCX files, extracts Client Name/ID (from the filename), and all dated case notes.
3.  **Convert:** Transforms the extracted data from CSV into multi-sheet, formatted Excel (.xlsx) workbooks.

### STEP 1: Input

Place your legacy client files (e.g., `archive.zip`) in the same directory as the scripts.

### STEP 2: Execute the Script

Run the following commands to make the shell script executable and start the process:
```bash
sudo chmod +x golive.sh
./golive.sh
```

The script will prompt you for two inputs:

1.  **Input ZIP File Name:** The name of your master archive (e.g., `archive.zip`).
2.  **Output Prefix:** A simple word to prefix the final output folder and temporary files (e.g., `batch_A`).

### STEP 3: OUTPUT AND CLEANUP


The process is self-cleaning, deleting temporary files upon completion.

Final Output Location:
A new directory will be created, named something like **`[Your Prefix]_xlsx_output`** (e.g., `batch_A_xlsx_output`).

### Output Contents:

This directory will contain one **multi-sheet Excel (.xlsx)** file for every client case file found. Each Excel file is structured as follows:

* **Sheet 1: Profile** (Key/Value pairs for Client/HMIS ID, formatted with bold fields and row banding. Case Managers will have to input current participants profile page manually. Only manual step of the conversion process).
* **Sheet 2: Case Notes** (Tabular data: Date, Staff, Note content).

============================================================


============================================================

