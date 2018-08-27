# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import docx
from win32com.client import Dispatch
import os



# print doc
OUTPUT_DIR = r'C:\MyProgram\doc'
word = Dispatch('word.application')
word.displayalerts = 0
word.visible = 0
countDoc = word.Documents.Count
print(countDoc)
doc = word.Documents.Open(r'C:\MyProgram\doc\test0.doc')
doc.SaveAs(os.path.join(OUTPUT_DIR, 'test0-new.docx'))
doc = docx.Document(r'C:\MyProgram\doc\test0-new.docx')
print doc
