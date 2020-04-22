#coding:UTF-8
from __future__ import division, print_function, absolute_import, unicode_literals


import sys
reload(sys) #不加这句容易报错
sys.setdefaultencoding('utf-8') #不加这句容易报错

import time
log_file = open('log_file.txt','a')
sys.stderr = log_file
sys.stdout = log_file
print(time.strftime("%Y-%m-%d %X") + '::Execute')

import main_kivy
