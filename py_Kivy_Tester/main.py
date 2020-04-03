#coding:UTF-8
from __future__ import division,print_function,absolute_import

import sys
import time
log_file = open('log_file.txt','a')
sys.stderr = log_file
sys.stdout = log_file
print(time.strftime("%Y-%m-%d %X") + '::Execute')

import main_kivy
