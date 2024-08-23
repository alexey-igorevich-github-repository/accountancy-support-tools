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

import sys
import os

import re

import time
import threading
import psutil

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QWidget, QPushButton, QListWidget, QLineEdit, QFileDialog
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QProgressBar, QSpacerItem, QSizePolicy, QScrollArea, QGridLayout

from PyQt6.QtCore import QThread, pyqtSignal
from multiprocessing import Manager

import logging
import logging.config
from logging_config import logger_conf

import miniaudio

from main import initialize_conversion, clear_com_cache

#############################################################

logging.config.dictConfig(logger_conf)
logger = logging.getLogger('the_logger')

#############################################################

class Monitor:
    def __init__(self, total_files):
        self.total_files = total_files
        self.start_time = time.time()
        self.processed_files = 0
        self.lock = threading.Lock()

    def update_progress(self, files_processed=1):
        with self.lock:
            self.processed_files += files_processed

    def get_statistics(self):
        elapsed_time = time.time() - self.start_time
        avg_speed = self.processed_files / elapsed_time if elapsed_time > 0 else 0
        remaining_files = self.total_files - self.processed_files
        estimated_time_left = remaining_files / avg_speed if avg_speed > 0 else float('inf')
        
        num_processes = len(psutil.Process().children(recursive=True))
        num_threads = threading.active_count()
        num_descriptors = len(psutil.Process().open_files())
        
        return {
            'avg_speed': avg_speed,
            'processed_files': self.processed_files,
            'remaining_files': remaining_files,
            'elapsed_time': elapsed_time,
            'estimated_time_left': estimated_time_left,
            'num_processes': num_processes,
            'num_threads': num_threads,
            'num_descriptors': num_descriptors
        }

#############################################################

class ConversionThread(QThread):
    def __init__(self,zipped_to_array_many_single_file_args,progress_queue):
        super().__init__()
        self.zipped_to_array_many_single_file_args = zipped_to_array_many_single_file_args
        self.progress_queue = progress_queue        

    def run(self):
        initialize_conversion(self.zipped_to_array_many_single_file_args, self.progress_queue)

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.selected_files = []     
        self.dest_folder = os.getcwd()       
        self.initUI()
        self.monitor = None
        self.progress_queue = None
        self.conversion_thread = None

    def initUI(self):
        self.setWindowTitle('Mass Pdf Converter')
        self.setWindowIcon(QIcon('static/images/icon.ico'))
        self.setGeometry(100, 100, 785, 1000)
        self.setFixedSize(785, 1000)
        self.setStyleSheet("background-color: black;")
        self.setWindowOpacity(0.95)
#------->
        main_layout = QVBoxLayout()
#------->
        self.listbox = QListWidget()
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(False)
        scroll_area.setWidget(self.listbox)
        main_layout.addWidget(self.listbox)        
#------->
        three_button_layout = QHBoxLayout()

        self.open_files_btn = QPushButton('Select Files')
        self.open_files_btn.setObjectName("open_files_btn")
        self.open_files_btn.clicked.connect(self.open_file)
        self.open_files_btn.setFixedSize(250, 50)
        three_button_layout.addWidget(self.open_files_btn)

        self.delete_selected_from_listbox_btn = QPushButton('Deselect')
        self.delete_selected_from_listbox_btn.setObjectName("delete_selected_from_listbox_btn")
        self.delete_selected_from_listbox_btn.clicked.connect(self.delete_selected_from_listbox)
        self.delete_selected_from_listbox_btn.setFixedSize(250, 50)
        three_button_layout.addWidget(self.delete_selected_from_listbox_btn)

        self.clear_listbox_btn = QPushButton('Clear Listbox')
        self.clear_listbox_btn.setObjectName("clear_listbox_btn")
        self.clear_listbox_btn.clicked.connect(self.clear_listbox)
        self.clear_listbox_btn.setFixedSize(250, 50)
        three_button_layout.addWidget(self.clear_listbox_btn)

        three_button_layout.addStretch()
  
        main_layout.addLayout(three_button_layout)
#------->   
        spacer1 = QSpacerItem(0, 25, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        main_layout.addSpacerItem(spacer1)

        self.dest_folder_entry = QLineEdit(self.dest_folder)
        main_layout.addWidget(self.dest_folder_entry)

        self.open_folder_btn = QPushButton('Select Destination Folder')
        self.open_folder_btn.setObjectName("open_folder_btn")
        self.open_folder_btn.clicked.connect(self.open_folder)
        self.open_folder_btn.setFixedSize(250, 50)
        main_layout.addWidget(self.open_folder_btn)

        spacer2 = QSpacerItem(0, 50, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        main_layout.addSpacerItem(spacer2)

        self.convert_btn = QPushButton('Convert')
        self.convert_btn.setObjectName("convert_btn")
        self.convert_btn.clicked.connect(self.convert)
        self.convert_btn.setFixedSize(250, 50)
        main_layout.addWidget(self.convert_btn)
        
        spacer3 = QSpacerItem(0, 25, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Fixed)
        main_layout.addSpacerItem(spacer3)       
 #------->       
        metrics_layout = QGridLayout()

        self.num_processes_label = QLabel("Processes: 0")
        metrics_layout.addWidget(self.num_processes_label, 0, 0)
        self.num_threads_label = QLabel("Threads: 0")
        metrics_layout.addWidget(self.num_threads_label, 1, 0)
        self.num_descriptors = QLabel("Descriptors: 0")
        metrics_layout.addWidget(self.num_descriptors, 2, 0)
        self.elapsed_time_label = QLabel("Elapsed Time: Infinite")
        metrics_layout.addWidget(self.elapsed_time_label, 0, 1)
        self.avg_speed_label = QLabel("Average Speed: 0 files/sec")
        metrics_layout.addWidget(self.avg_speed_label, 1, 1)
        self.estimated_time_label = QLabel("Estimated Time Left: Infinite")
        metrics_layout.addWidget(self.estimated_time_label, 2, 1)
        self.processed_files_label = QLabel("Processed Files: 0")
        metrics_layout.addWidget(self.processed_files_label, 0, 2)        
        self.remaining_files_label = QLabel("Remaining Files: 0")
        metrics_layout.addWidget(self.remaining_files_label, 1, 2)
        
        main_layout.addLayout(metrics_layout)
#------->       
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setMaximum(100)
        main_layout.addWidget(self.progress_bar)
#------->
        self.setLayout(main_layout)

    def open_file(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Open Files')
        if files:
            self.selected_files.extend(files)
            self.update_listbox()
    
    def delete_selected_from_listbox(self):
        selected_items = self.listbox.selectedItems()
        if not selected_items:
            return     
        for item in selected_items:
            file_path = item.text()
            if file_path in self.selected_files:
                self.selected_files.remove(file_path)
        self.update_listbox()

    def update_listbox(self):
        self.listbox.clear()
        self.listbox.addItems(self.selected_files)

    def clear_listbox(self):
        self.selected_files.clear()
        self.listbox.clear()
        self.listbox.addItems(self.selected_files)

    def open_folder(self):
        self.dest_folder = QFileDialog.getExistingDirectory(self, 'Select Destination Folder')
        self.dest_folder_entry.setText(self.dest_folder)
        print(self.dest_folder)

    def extract_input_path_and_input_file(self, selected_file):
        pattern = r"^(.*[\\/])([^\\/]+)$"
        match = re.match(pattern, selected_file)
        if match:
            input_path = match.group(1)
            input_file = match.group(2)
            print(f'{selected_file} == {input_path}, {input_file}')
            return input_path, input_file
        else:
            print(f'{selected_file} == None, None')
            return None, None

    def convert(self):                
        self.progress_bar.setValue(int(0))

        if not self.selected_files:
            return
        
        input_files = []
        input_paths = []
        output_paths = []

        for selected_file in self.selected_files:
            print(selected_file)
            input_path, input_file = self.extract_input_path_and_input_file(selected_file)
            input_paths.append(input_path)
            input_files.append(input_file)

        output_paths = [self.dest_folder] * len(input_files)
        zipped_to_array_many_single_file_args = list(zip(input_paths, output_paths, input_files))
        self.monitor =  Monitor(len(zipped_to_array_many_single_file_args))    

        manager = Manager()  
        self.progress_queue = manager.Queue() 
        self.conversion_thread = MonitorThread(self.progress_queue, len(zipped_to_array_many_single_file_args))  
        self.conversion_thread.update_signal.connect(self.update_progress)   
        self.conversion_thread.finished.connect(self.on_conversion_finished)   
        self.conversion_thread.start() 

        clear_com_cache()

        logger.debug(f"Input files: {input_files}")
        
        self.worker_thread = ConversionThread(zipped_to_array_many_single_file_args, self.progress_queue)
        self.worker_thread.start()
   

    def update_progress(self, processed, total):      
        self.monitor.update_progress(files_processed = processed - self.monitor.processed_files)
        stats = self.monitor.get_statistics()

        self.avg_speed_label.setText(f"Average Speed: {stats['avg_speed']:.2f} files/sec")
        self.processed_files_label.setText(f"Processed Files: {stats['processed_files']}")
        self.remaining_files_label.setText(f"Remaining Files: {stats['remaining_files']}")
        self.estimated_time_label.setText(f"Estimated Time Left: {stats['estimated_time_left']:.2f} sec")
        self.elapsed_time_label.setText(f"Elapsed Time: {stats['elapsed_time']:.2f} sec")
        self.num_processes_label.setText(f"Processes: {stats['num_processes']}")
        self.num_threads_label.setText(f"Threads: {stats['num_threads']}")
        self.num_descriptors.setText(f"Descriptors: {stats['num_descriptors']}")

        progress = int((processed / total) * 100)
        self.progress_bar.setValue(progress)


    def on_conversion_finished(self):
        stream = miniaudio.stream_file("static/sounds/ready.wav")
        device = miniaudio.PlaybackDevice()
        device.start(stream)

##########################################################################

class MonitorThread(QThread):
    update_signal = pyqtSignal(int, int)

    def __init__(self, progress_queue, total_files):
        super().__init__()
        self.progress_queue = progress_queue
        self.total_files = total_files
        self.processed_files = 0
    
    def run(self):
        while self.processed_files < self.total_files:
            if not self.progress_queue.empty():
                _, file = self.progress_queue.get()
                self.processed_files += 1
                self.update_signal.emit(self.processed_files, self.total_files)
            time.sleep(0.1)  

##########################################################################

if __name__ == '__main__':

    app = QApplication(sys.argv)

    with open("static/styles/style.qss", "r") as file:
        app.setStyleSheet(file.read())

    ex = App()
    ex.show()
    sys.exit(app.exec())

########################################################################## 

    
