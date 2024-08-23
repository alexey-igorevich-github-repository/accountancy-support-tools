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

#############################################################

import os
import sys

import re
import shutil

import multiprocessing

from multiprocessing import cpu_count

from multiprocessing import current_process

from concurrent.futures import ProcessPoolExecutor


import pythoncom

import win32com.client

import comtypes.client  

from fpdf import FPDF

from PIL import Image

import time

import logging
import logging.config
from logging_config import logger_conf


import pyodbc
import pandas as pd

from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, PageBreak

import pdfkit

from bs4 import BeautifulSoup

#############################################################

def clear_com_cache():
    folder = win32com.client.gencache.GetGeneratePath()
    if os.path.exists(folder):
        shutil.rmtree(folder)

#############################################################

logging.config.dictConfig(logger_conf)
logger = logging.getLogger('my_python_logger')

#############################################################

multiprocessing.set_start_method('spawn', force=True)

#############################################################

def convert_xlsx_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Excel to PDF)")
    pythoncom.CoInitialize()
    try:    
        excel_com_object = comtypes.client.CreateObject("Excel.Application")
        excel_com_object.Visible = False
        workbook = excel_com_object.Workbooks.Open(input_filename)  
        workbook.ExportAsFixedFormat(0, output_filename)
        workbook.Close(SaveChanges=False)
        logger.info(f"Successfully converted {input_filename} to {output_filename}")
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    finally:        
        if 'workbook' in locals():
            workbook.Close()
        if 'excel_com_object' in locals():
            excel_com_object.Quit()
            del excel_com_object
        pythoncom.CoUninitialize()

#############################################################

def convert_docx_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Word to PDF)")    
    pythoncom.CoInitialize()
    try:
        word_com_object = comtypes.client.CreateObject("Word.Application") 
        word_com_object.Visible = False
        document = word_com_object.Documents.Open(input_filename)
        document.SaveAs(output_filename, FileFormat=17) 
        document.Close(SaveChanges=False)
        logger.info(f"Successfully converted {input_filename} to {output_filename}")
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    finally:
        if 'document' in locals():
            document.Close()
        if 'word_com_object' in locals():
            word_com_object.Quit()
            del word_com_object 
        pythoncom.CoUninitialize()

#############################################################

def convert_pptx_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (PowerPoint to PDF)")    
    pythoncom.CoInitialize()
    try:
        powerpoint_com_object = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint_com_object.Visible = True
        powerpoint_com_object.WindowState = 2 
        presentation = powerpoint_com_object.Presentations.Open(input_filename)
        presentation.SaveAs(output_filename, 32)
        presentation.Close(SaveChanges=False)  
        logger.info(f"Successfully converted {input_filename} to {output_filename}")
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    finally:
        if 'presentation' in locals():
            presentation.Close()
        if 'powerpoint_com_object' in locals():
            powerpoint_com_object.Quit()
            del powerpoint_com_object
        pythoncom.CoUninitialize()

#############################################################

def convert_pub_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Publisher to PDF)")    
    pythoncom.CoInitialize()   
    try:
        publisher_com_object = comtypes.client.CreateObject("Publisher.Application")        
        publication = publisher_com_object.Open(input_filename)
        publication.ExportAsFixedFormat(2, output_filename)      
        logger.info(f"Successfully converted {input_filename} to {output_filename}")
        publication.Close()
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    finally:
        if 'publication' in locals():
            publication.Close()
        if 'publisher_com_object' in locals():
            publisher_com_object.Quit()
            del publisher_com_object 
        pythoncom.CoUninitialize() 

#############################################################

def convert_txt_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Text to PDF)")
    try:    
        pdf = FPDF()
        pdf.add_page()
        pdf.set_margins(left=10, top=15, right=10)
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.set_font("Times", '', 14)

        page_width = pdf.w - pdf.l_margin - pdf.r_margin
        cell_width = page_width
        cell_height = pdf.font_size * 1.5

        with open(input_filename, 'r', encoding='utf-8') as file:
            for line in file:
                pdf.multi_cell(cell_width, cell_height, txt=line.strip())
        
        pdf.output(output_filename)
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    

#############################################################

def convert_img_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Image to PDF)")
    try:
        image = handle_transparency(input_filename)        
        image.save(output_filename)
        logger.info(f"Successfully converted {input_filename} to {output_filename}")
    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)

#------------------------------------------------------------

def handle_transparency(ifc): 
    ifc = Image.open(ifc)
    if ifc.mode in ("RGBA", "LA") or (ifc.mode == "P" and "transparency" in ifc.info):
        ifc.load()
        background = Image.new("RGB", ifc.size, (255, 255, 255))
        background.paste(ifc, mask=ifc.split()[3])
        ifc = background
    else:
        ifc = ifc.convert("RGB")
    return ifc


#############################################################

def convert_access_to_pdf(input_filename, output_filename):
    logger.debug(f"Starting conversion of {input_filename} to {output_filename} (Access to PDF)")
    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={input_filename};'
        )

        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        tables = [row.table_name for row in cursor.tables(tableType='TABLE')]

        pdf_filename = f'{output_filename}'
        doc = SimpleDocTemplate(pdf_filename,
            pagesize=landscape(A4),
            leftmargin=20,
            rightmargin=20,
            topmargin=40,
            bottommargin=20
        )


        elements = []


        def split_dataframe_by_columns(df, max_columns_per_page):
            columns = df.columns
            num_columns = len(columns)
            chunks = [df.iloc[:, i:i + max_columns_per_page] for i in range(0, num_columns, max_columns_per_page)]
            return chunks

        def create_table(df_chunk):
            data = [df_chunk.columns.tolist()] + df_chunk.values.tolist()
            t = Table(data, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ]))
            return t

        for table in tables:
            query = f"SELECT * FROM {table}"
            df = pd.read_sql(query, conn)            
            max_columns_per_page = 10
            df_chunks = split_dataframe_by_columns(df, max_columns_per_page)            
            for df_chunk in df_chunks:
                table = create_table(df_chunk)
                elements.append(table)
                elements.append(PageBreak())
            elements.append(Spacer(1, 12))
        conn.close()
        doc.build(elements)

    except Exception as e:
        logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
    finally:
        try:
            conn.close()
        except:
            pass

#############################################################

def convert_web_page_to_pdf(input_filename, output_filename):     
    path_to_wkhtmltopdf = "modules\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"  

    config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
    options_for_local_file = {
        'no-outline': None,
        'disable-local-file-access': None,
        'enable-local-file-access': None,
        'load-error-handling': 'ignore',
        'load-media-error-handling': 'ignore',
    }

    if input_filename.endswith(("_bookmarks.html", "_links.html")):
        logger.debug(f"Starting conversion of {input_filename} which appears to be a URL list to many pdfs (Web page to PDF)")
        url_array = extract_urls_from_browser_bookmarks(input_filename)
        digest_url_array(url_array, input_filename, output_filename)

    elif input_filename.endswith(("_bookmarks.txt", "_links.txt")):
        logger.debug(f"Starting conversion of {input_filename} which appears to be a URL list to many pdfs (Web page to PDF)")
        url_array = read_urls_from_plaintext_file(input_filename)
        digest_url_array(url_array, input_filename, output_filename)
    
    else:
        input_file = input_filename
        output_file = output_filename
        try:
            logger.debug(f"Starting conversion of {input_filename} to {output_filename} (HTML to PDF)")
            pdfkit.from_file(input_file, output_file, configuration=config, options=options_for_local_file)
            logger.info(f"Successfully converted {input_filename} to {output_filename}")
        except Exception as e:
            logger.error(f"Failed to convert {input_filename} to PDF: {e}", exc_info=True)
        finally:
            pass

#-------------------------------------------------------------------------

def extract_urls_from_browser_bookmarks(file_path):
    url_array = []
    
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')
        
        for a_tag in soup.find_all('a'):
            href = a_tag.get('href')
            if href:
                url_array.append(href)
    
    return url_array

#-------------------------------------------------------------------------

def read_urls_from_plaintext_file(file_path):
    url_array = []
    
    with open(file_path, 'r', encoding='utf-8') as file:
        for line in file:
            url = line.strip()
            if url:
                url_array.append(url)
    
    return url_array

#-------------------------------------------------------------------------

def get_directory_path(absolute_file_path):
    last_slash_index = max(absolute_file_path.rfind('/'), absolute_file_path.rfind('\\'))
    if last_slash_index == -1:
        return ''
    directory_path = absolute_file_path[:last_slash_index]
    return directory_path

#-------------------------------------------------------------------------

def generate_filename_from_url(url, output_filename):
    output_directory = get_directory_path(output_filename)
    output_file = url.replace(':', '')
    output_file = re.sub(r'[<>:"/\\|?*]', '-', output_file)
    output_file = f"{output_directory}\\{output_file}.pdf"
    return output_file

#-------------------------------------------------------------------------

def digest_url_array(url_array, input_filename, output_filename):
    path_to_wkhtmltopdf = "modules\\wkhtmltopdf\\bin\\wkhtmltopdf.exe" 

    config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)

    try:
        for url in url_array:

            output_file = generate_filename_from_url(url, output_filename)
            logger.debug(f"Starting conversion of {url} to {output_file} (URL to PDF)")

            try:
                pdfkit.from_url(url, output_file, configuration=config)
                logger.info(f"Successfully converted {url} to {output_file}")
            except Exception as e:
                logger.error(f"Failed to convert {url} to {output_file}: {e}", exc_info=True)
            finally:
                pass    

    except Exception as e:
        logger.error(f"Failed to convert {input_filename}  which appears to be a URL list to many pdfs: {e}", exc_info=True)

    finally:
        pass

#############################################################

def generate_input_filename(input_path, filename):
    input_filename = os.path.join(input_path, filename) 
    return input_filename


def generate_output_filename(output_path, input_filename):
    output_filename = os.path.split(input_filename)[-1] 
    output_filename = os.path.splitext(output_filename) 
    output_filename = f"{output_filename[0]}.pdf"
    output_filename = os.path.join(output_path, output_filename)
    output_filename = output_filename.replace("/", "\\") 
    return output_filename

##########################################################################

def process_file(single_file_args):
    input_path, output_path, filename = single_file_args
    try:
        convert_file(input_path, output_path, filename)
    except Exception as e:
        logger.error(f"Error processing file {filename}: {e}", exc_info=True)

##########################################################################

def convert_file(input_path, output_path, filename):

    input_filename = generate_input_filename(input_path, filename)

    if input_filename.endswith(("_bookmarks.txt", "_links.txt", "_bookmarks.html", "_links.html", ".html", ".htm")):
        convert_web_page_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith(".xlsx") or input_filename.endswith(".xls"):       
        convert_xlsx_to_pdf((input_filename), generate_output_filename(output_path, input_filename))
     
    elif input_filename.endswith(".docx") or input_filename.endswith(".doc"):     
        convert_docx_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith(".pptx") or input_filename.endswith(".ppt"):
        convert_pptx_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith(".txt"):
        convert_txt_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith((".blp", ".blp1", ".blp2", ".bmp", ".dds", ".dxt1", ".dxt5", ".dib", ".eps", ".gif", ".icns", ".ico", ".im", ".jp2", ".jpx", ".j2k", ".msp", ".pcx", ".pfm", ".png", ".apng", ".ppm", ".sgi", ".spider", ".tga", ".tiff", ".webp", ".xbm", ".cur", ".dcx", ".fits", ".fli", ".flc", ".fpx", ".ftex", ".gbr", ".gd", ".gd2", ".imt", ".iptc", ".naa", ".mcidas", ".mic", ".mpo", ".pcd", ".pixar", ".qoi", ".sun", ".wal", ".wmf", ".emf", ".xpm", ".jpg", ".jpeg")):
        convert_img_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith((".accdb", ".mdb")):
        convert_access_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    elif input_filename.endswith(".pub"):
        convert_pub_to_pdf((input_filename), generate_output_filename(output_path, input_filename))

    else:
        logger.warning(f"Unsupported file type for {filename}")
   
##########################################################################

def calculate_usable_cpu():

    cpu_number = os.cpu_count()

    match cpu_number:
        case 1:
            available_cpu = 1
        case 2:
            available_cpu = 1
        case 3:
            available_cpu = 2
        case 4:
            available_cpu = 3
        case 5:
            available_cpu = 3
        case 6:
            available_cpu = 4
        case _:
            available_cpu = cpu_number - 3

    return available_cpu

##########################################################################

def worker(single_file_args, progress_queue):
    process_file(single_file_args)
    progress_queue.put((current_process().pid, single_file_args[2]))

##########################################################################        

def initialize_conversion(zipped_to_array_many_single_file_args, progress_queue=None):

#-------------------------------------------------------------------------

    start_time = time.perf_counter()
    logger.debug("Script started")

#-------------------------------------------------------------------------

    with ProcessPoolExecutor(max_workers=calculate_usable_cpu(),max_tasks_per_child=4) as executor: 
        futures = [executor.submit(worker, single_file_args, progress_queue) for single_file_args in zipped_to_array_many_single_file_args] 

        for future in futures:
            future.result()

#-------------------------------------------------------------------------        
    end_time = time.perf_counter()
    total_execution_time = {end_time - start_time}
    logger.info(f"Script completed in {end_time - start_time:.2f} seconds")

##########################################################################