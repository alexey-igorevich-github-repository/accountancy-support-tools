# excel-to-word-report-generator(pywin32)

--------------------------------------------------------------------

This project is a Python-based tool for generating Word reports from data stored in Excel spreadsheets. The tool allows for dynamic text replacement in Word templates, making it a versatile solution for creating personalized documents.

--------------------------------------------------------------------
## Table of Contents
0. [SuperfastHowTo](#SuperfastHowTo)
1. [Features](#features)
2. [Requirements](#requirements)
3. [Installation](#installation)
4. [Usage](#usage)
5. [Directory Structure](#directory-structure)
6. [Error Handling](#error-handling)
7. [License](#license)
8. [Support](#support)

--------------------------------------------------------------------

### SuperfastHowTo

- download
- run `excel-to-word-report-generator(pywin32).exe`
- go have a look in `/output`
- go to `/input` and make your corrections
- run again
- go to `/output` take your files

--------------------------------------------------------------------
### Features

- **Dynamic Text Replacement**: Replace placeholder text in Word documents with data from Excel.

- **Supports Multiple Elements**: Works with paragraphs, tables, text boxes, shapes, and drawings in Word.

- **Error Handling**: Displays error messages using a GUI if any issues occur during processing.

- **Automated Report Generation**: Generates reports with unique filenames based on current datetime and data from Excel.

--------------------------------------------------------------------  

### Requirements

To run this project, you will needto install some libraries:

```bash
pip install -r requirements.txt
```

--------------------------------------------------------------------

### Installation

> [!Tip] This version only for Windows users (you will need Word and Excel):
1. Download `excel-to-word-report-generator(pywin32).exe` file, place it in some folder
2. First launch `excel-to-word-report-generator(pywin32).exe` file to create the required structure

--------------------------------------------------------------------

### Usage
##### 1. Prepare Your Data:
- Inside "table.xlsx" you can rename your "human_sheet" but you shouldn't rename the "machine_sheet".

- Inside "table.xlsx", on the "human_sheet" you draw a table with the structure you like. This is for you to have a convinient workspace for frequently usage.

- Inside "table.xlsx", inside the "machine_sheet" make links to the cells inside the "human_sheet" (You can make links to whereever you want). Just click the cell, tap = and click the required cell. You can then drag a frame to quickly copy the formula and so on. Obvious Excel working process.

- Ensure
    - the directory structure is correct
    - the Excel file is named table.xlsx
    - the Excel file has a sheet named machine_sheet.
    - "machine_sheet" has a proper structure:
        - First row is for human-readable definitions
        - Second row is for placeholders {Something} format
        - Third row and lower are for your data
    - the Word file is named "template.docx"
    - the Word file has something to replace for you to see the result
##### 2. Run the Script:
excel-to-word-report-generator(pywin32).exe

##### 3. Take your reports from /output.

  --------------------------------------------------------------------
  
### Directory Structure

```bash
excel-to-word-report-generator(pywin32)/
│
├── input/
│ ├── table.xlsx # Here you put your data
│ ├── template.docx # Here you make your skeleton-form for quick data injection
│ └── README.md # Info about the program
│
├── output/ # Here you take your reports
│
├── templates/ # From here the skeleton-files are copied to /input if not exist
│ ├── table.xlsx
│ ├── template.docx
│ └── README.md
│
├── excel-to-word-report-generator(pywin32).py # Run it to generate your reports
└── requirements.txt # What is required to install to have the program working
```


--------------------------------------------------------------------

### Error Handling

The script is able to handle exceptions gracefully. If an error occurs, a small window with error message will be displayed. The problem-function execution will be skipped. Other functions will continue their work. So remember that if you see an error message - it means that some of your placeholders were not replaced with data.

--------------------------------------------------------------------
  
### License

```
# This file is part of [accountancy-support-tools]. 
# # [accountancy-support-tools] is free software: you can redistribute it and/or 
# modify it under the terms of the GNU General Public License as published by 
# the Free Software Foundation, either version 3 of the License, or (at your option) 
# any later version. 
# # [accountancy-support-tools] is distributed in the hope that it will be useful, 
# but WITHOUT ANY WARRANTY; without even the implied warranty of 
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the 
# GNU General Public License for more details. 
# # You should have received a copy of the GNU General Public License 
# along with [accountancy-support-tools]. If not, see <http://www.gnu.org/licenses/>.
```
  
--------------------------------------------------------------------
  
### Support

> [!Success] Your tips are always welcome XD


 BTC: bc1qfkuxgu0vkl5u5pr2l0uag74a4u2273w2a9j95f <br>
<div><p><a href="https://yoomoney.ru/to/4100118693354177"> <img align="left" src="https://avatars.githubusercontent.com/u/6553002?s=200&v=4" height="75" width="75" alt="Yoomoney" /></a></p></div>

<div><p><a href="https://ko-fi.com/alexey_i_c"> <img align="left" src="https://cdn.ko-fi.com/cdn/kofi3.png?v=3" height="50" width="210" alt="Buy me a coffee" /></a></p></div>


--------------------------------------------------------------------

© 2024 Alexey Igorevich. All rights reserved.