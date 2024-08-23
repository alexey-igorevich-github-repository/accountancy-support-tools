# mass-pdf-converter

--------------------------------------------------------------------

This project is a Python-based tool for converting multiple files to pdfs. Also it can parse your bookmarks export from browser and make a pdf snapshot of every page listed. This tool uses multiprocessing and has a gui.

--------------------------------------------------------------------
## Table of Contents
1. [SuperfastHowTo](#SuperfastHowTo)
2. [Features](#Features)
3. [Requirements](#Requirements)
4. [Installation](#Installation)
5. [Usage](#Usage)
6. [Troubleshooting](#Troubleshooting)
7. [Structure](#Structure)
8. [Error Handling](#Error-handling)
9. [License](#License)
10. [Special Thanks](#Special-thanks)
11. [Support](#Support)

--------------------------------------------------------------------

### SuperfastHowTo

- download
- install dependencies `requirements.txt`
- install `wkhtmltopdf` and put it into `modules` directory
- run `gui.py`
- choose files to convert
- choose output folder
- click convert
- go to your output folder and have your result

--------------------------------------------------------------------

#### To parse links that are inside your html file (exported browser bookmarks)
 Rename your exported browser bookmarks files to have one of this endings:
 - _bookmarks.html 
 - _links.html
 - _bookmarks.txt
 - _links.txt

--------------------------------------------------------------------
### Features

- **Gui**: Has a convinient Gui for you to have a visualization of the process.

- **Multiprocessing**: Utilizes multi-threading and multiprocessing to speed up conversions, taking advantage of all available CPU cores. Faster conversion of a big mass of files or web-pages. Auto calculation of cpu power.

- **Progress Monitoring**: Real-time display of conversion progress, including metrics like average speed, number of files processed, and estimated time remaining.

- **Logging**: Comprehensive logging capabilities to track the conversion process, including detailed error handling and debug information.

- **Supported Formats**: Excel (.xlsx, .xls), Word (.docx, .doc), PowerPoint (.pptx, .ppt), Publisher (.pub), Images (.jpg, .png, .bmp, and more), HTML files, and even Microsoft Access databases (.accdb, .mdb).

--------------------------------------------------------------------  

### Requirements

Before running the application, ensure you have the following installed:

- Python 3.12 `python.org`
- requirements.txt `pip install -r requirements.txt`
- wkhtmltopdf `https://wkhtmltopdf.org/` place into `mass-pdf-converter/modules/wkhtmltopdf`

--------------------------------------------------------------------

### Installation

> [!Tip] Application was designed for Windows users:
1. You need Python 3.12 installed
2. Download the zip, unzip, open folder with something like `Visual Studio Code`.
3. You need to make a virtual environment `python -m venv venv` then activate it `\venv'Scripts\activate`
4. You need dependencies to be installed `pip install -r requirements.txt`
5. The `wkhtmltopdf` should be downloaded from here `https://wkhtmltopdf.org/` and placed into `modules` directory
6. Now you can run `gui.py` 

--------------------------------------------------------------------

### Usage
##### 0. Running the Application:
    Run "gui.py" file in the dist directory.

##### 1. User Interface:
    File Selection:
        Use the "Select Files" button to choose the files you wish to convert.
        The selected files will appear in the listbox.
        You can remove files from the list by selecting them and clicking "Deselect" or clear the entire list using "Clear Listbox".

    Destination Folder:
        Set the destination folder where the converted PDF files will be saved using the "Select Destination Folder" button.
        The path will appear in the destination folder entry field.

    Conversion:
        Click the "Convert" button to start the batch conversion process.
        The progress bar and statistics will update in real-time, showing the number of files processed, average speed, and estimated time remaining.

##### 2. Logging:
    The application logs all activities, including errors and debug information, into a logs directory.
    Log files are categorized by severity: notset.log, debug.log, info.log, warning.log, error.log, critical.log.

##### 3. Supported File Formats
    Documents:
        Microsoft Word: .docx, .doc
        Microsoft Excel: .xlsx, .xls
        Microsoft PowerPoint: .pptx, .ppt
        Microsoft Publisher: .pub
        Microsoft Access: .accdb, .mdb
        Plain Text: .txt
    Images: .jpg, .jpeg, .png, .bmp, .gif, .tiff, and more.
    Web Files: .html, .htm
    Other files for parsing: exported from browsers HTML bookmark files, TXT files containing URLs 
    (must end with: _bookmarks.html || _links.html || _bookmarks.txt || _links.txt).

  --------------------------------------------------------------------
  
### Troubleshooting
Common Issues

    Missing Dependencies: Ensure all necessary Python packages are installed. Run `pip install -r requirements.txt` to resolve missing dependencies.
    File Access Issues: Make sure the application has read/write access to the files and directories involved in the conversion process.
    Conversion Errors: Check the log files in the logs directory for detailed error messages.

Known Limitations

    Large Files: Very large files may take considerable time to convert, especially when dealing with complex Excel workbooks or PowerPoint presentations.
    File Types: Not all file types are supported for conversion. Unsupported files will be logged and skipped.


### Structure

```bash
mass-pdf-converter/ 
│
├── static/ 
├── modules/wkhtmltopdf # You can take it from Program Files
├── logs/ # Here you can see logging info
│
├── main.py
├── gui.py # Run it to show you a convinient gui window
├── logging_config.py
├── requirements.txt # Used packages
└── README.md # This readme file you are reading now
```

--------------------------------------------------------------------

### Error Handling

The script is able to handle exceptions gracefully. 
And it writes logs. If an error occurs, you will see nothing. XD
Check /logs to see if everything has been converted.

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

### Special Thanks:
My utmost respect for hard working development teams and solo professionals 
who have provided "the tools and construction blocks written on low-level 
languages" which made possible so easy assemblement of this application.

Thanks `Suraj Singh` `suraj-singh12` for an inspiration example: https://github.com/suraj-singh12/any-to-pdf

Thanks the `Floraphonic` audio-stock shop and the guy who has created the `cute-level-up.mp3` sound-effect for providing it free: https://www.floraphonic.com/

Thanks to `David Faure` and `KDABLabs` `https://www.kdab.com` for providing an example of .qss styles: https://github.com/KDABLabs/kdabtv/tree/master/Styling-Qt-Widgets

--------------------------------------------------------------------
  
### Support

> [!Success] Your tip is always welcome XD


 BTC: bc1qfkuxgu0vkl5u5pr2l0uag74a4u2273w2a9j95f <br>
<div><p><a href="https://yoomoney.ru/to/4100118693354177"> <img align="left" src="https://avatars.githubusercontent.com/u/6553002?s=200&v=4" height="75" width="75" alt="Yoomoney" /></a></p></div>

<div><p><a href="https://ko-fi.com/alexey_i_c"> <img align="left" src="https://cdn.ko-fi.com/cdn/kofi3.png?v=3" height="50" width="210" alt="Buy me a coffee" /></a></p></div>


--------------------------------------------------------------------

© 2024 Alexey Igorevich. All rights reserved.