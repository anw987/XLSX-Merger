# Excel Report Generator
This project is a Python-based tool to read multiple Excel files, transform and format the data, and generate a consolidated Excel report.

## Overview
The generated Excel file includes:
*Multiple sheets
Custom headers and layout
Styled tables (colors, alternating rows)
Charts (bar and pie)
Auto-sized columns
Timestamp-based filename

Output file format:
Report-yymmddhhmm.xlsx

Libraries Used: 
##1. Pandas

Used for:
Reading Excel files
Data transformation and filtering
Column reordering
Pivoting data for charts

Install:
```bash
pip install pandas
```

##2. Openpyxl

Used for:
Writing Excel files
Styling cells (font, color, alignment)
Adjusting column widths
Creating charts (bar, pie)
Positioning elements inside sheets

Install:
```bash
pip install openpyxl
```

##3. datetime (built-in)

Used for:
Generating dynamic filename with timestamp

No installation required.

##4. PyInstaller

Used for:
Converting the Python script into a standalone executable (.exe)

Install:
```bash
pip install pyinstaller
```

##Project Preparation
Install Python (recommended version: 3.9 or higher)
2. Install required libraries:
```bash
pip install pandas openpyxl pyinstaller
```
3. Place all required input Excel files in the same directory as the script.
4. Ensure your main script (for example, merge.py) is in the same folder.
5. Running the Script

Execute the script using:
```bash
python merge.py
```

After execution, a new Excel file will be generated:
Report-<timestamp>.xlsx

##Features
Merge multiple Excel files into one workbook
Custom header insertion and layout control
Column reordering and data manipulation
Constant value injection into columns
Alternating row color styling
Auto-adjust column width based on content
Bar chart for time-series data
Pie chart for summary metrics
Dynamic chart positioning
Dashboard-style formatting
Building the Executable (.exe)

Basic build
```bash
py -m PyInstaller --onefile merge.py
```

Output will be available in:
dist/merge.exe

Build without console window
```bash
py -m PyInstaller --onefile --noconsole merge.py
```

Add version information (optional)
Create a file named version.txt containing version metadata
Build using:
```bash
py -m PyInstaller --onefile --noconsole --version-file=version.txt merge.py
```

Add custom icon (optional)
```bash
py -m PyInstaller --onefile --noconsole --icon=icon.ico merge.py
```

Running the EXE
Navigate to the dist folder
Run merge.exe

No Python installation is required on the target machine.

##Notes
Ensure input file names match exactly
Avoid opening the output file while running the script
Processing time depends on file size
Chart rendering may vary slightly across Excel versions
Possible Enhancements
Add configuration file for flexible input paths
Add logging mechanism
Support dynamic data structures
Improve dashboard layout and styling
Add CLI parameters for customization
