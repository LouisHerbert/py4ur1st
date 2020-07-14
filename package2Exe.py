# -*- coding: utf-8 -*-
'''
@author: wwang14
@date: Nov 27, 2019
'''
import os
# if no Pyinstaller, cmd 'pip install pyinstaller' pls.
# cmd = 'Pyinstaller -F -w main.py -i valeo.ico --hidden-import data.py --hidden-import office.py --hidden-import tsxml.py -n EXCEL2TSXML'
cmd = 'Pyinstaller --onefile Double_click_HELLO.py'

os.system(cmd)