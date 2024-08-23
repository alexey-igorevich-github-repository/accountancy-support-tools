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
import logging
from logging import LogRecord
logging.StreamHandler

############################################################################################################

os.makedirs('logs', exist_ok=True)

############################################################################################################

class ConsoleFilter(logging.Filter):
    def filter(self, record):
        if hasattr(record, 'status_code') and record.status_code != 200:
            print('warning http status is not OK')        
        return True
  
class NotsetFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return True 

class DebugFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return record.levelno == logging.DEBUG 
            
class InfoFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return record.levelno == logging.INFO
            
class WarningFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return record.levelno == logging.WARNING
            
class ErrorFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return record.levelno == logging.ERROR
            
class CriticalFilter(logging.Filter):
    def filter(self, record: LogRecord) -> bool:
        return record.levelno == logging.CRITICAL
           
############################################################################################################

logger_conf = {
    'version': 1,
    'formatters': {
        'console_msg_formatter': {
            'format': '{levelname} {msg} {filename} {funcName} {lineno} {exc_info}',
            'style': "{"
        },
    },
    'handlers': {
        'console': {                       
            'class': 'logging.StreamHandler',
            'level': 'NOTSET',
            'formatter': 'console_msg_formatter'
        },
        'file_notset': {                       
            'class': 'logging.FileHandler',
            'level': 'NOTSET',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/notset.log',
            'filters': ['notset_filter']            
        },        
        'file_debug': {                        
            'class': 'logging.FileHandler',
            'level': 'DEBUG',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/debug.log',
            'filters': ['debug_filter']             
        },
        'file_info': {
            'class': 'logging.FileHandler',
            'level': 'INFO',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/info.log',
            'filters': ['info_filter'] 
        },
        'file_warning': {
            'class': 'logging.FileHandler',
            'level': 'WARNING',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/warning.log',
            'filters': ['warning_filter'] 
        },
        'file_error': {
            'class': 'logging.FileHandler',
            'level': 'ERROR',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/error.log',
            'filters': ['error_filter'] 
        },
        'file_critical': {
            'class': 'logging.FileHandler',
            'level': 'CRITICAL',
            'formatter': 'console_msg_formatter',
            'filename': 'logs/critical.log',
            'filters': ['critical_filter'] 
        }
    },
    'filters': {
        'console_filter': {
            '()': ConsoleFilter     
        },
        'notset_filter': {
            '()': NotsetFilter     
        },
        'debug_filter': {
            '()': DebugFilter    
        },
        'info_filter': {
            '()': InfoFilter     
        },
        'warning_filter': {
            '()': WarningFilter     
        },
        'error_filter': {
            '()': ErrorFilter   
        },
        'critical_filter': {
            '()': CriticalFilter   
        }                                                            
    },
    'loggers': {
        'the_logger': {
            'level': 'DEBUG',
            'handlers': ['console', 'file_notset', 'file_debug', 'file_info', 'file_warning', 'file_error', 'file_critical'],   
            'filters': []
        }
    }
}

############################################################################################################