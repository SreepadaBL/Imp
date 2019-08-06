import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from win32com.client import constants
from collections import Counter
import docx2txt

count=0
##Read the path 
import sys                                                                                 
iter=sys.argv[1]
iter1=sys.argv[2]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 33. Occurrences of 'Shall', 'Will', 'may', 'bcos', '+ve', '-ve'., etc in Document.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
res=[]
l=[]
ll=[]
count=0
##Open the Document
if iter.endswith('.doc'):# or iter.endswith('.docx'):
 word1 =  win32.gencache.EnsureDispatch ("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Paragraphs
 word_list=iter1.split(',')
 if len(word_list)>=1:
  for p in para:
   ll=str(p.Range.Text.encode('ascii','ignore').decode()).split()
   for value in ll:
    value=value.strip('\r')
    value=value.strip('\r\x07')
    value=value.strip('\x0c')
    value=value.strip('\x0b')
    value=value.strip('\x0a')
    value=value.rstrip(' ')
    value=value.strip('\n')
    for i in word_list:
     if value == i:
      l.append(value)
  cnt=dict(Counter(l))
  for key, value in sorted(cnt.items(), key=lambda item: (item[0], item[1])):
   print ("Frequency of '"'%s'"': %s" %(key, value))  
 sheet_1.Close()
 word1.Quit()
elif iter.endswith('.docx'):
 doc1 = docx2txt.process(iter)
 para=doc1.splitlines()
 #print(para)
 word_list=iter1.split(',')
 if len(word_list)>=1:
  for p in para:
   if p!='':
    ll=p.encode('ascii','ignore').decode().split()
    for value in ll:
     value=value.strip('\r')
     value=value.strip('\r\x07')
     value=value.strip('\x0c')
     value=value.strip('\x0b')
     value=value.strip('\x0a')
     value=value.rstrip(' ')
     value=value.strip('\n')
     for i in word_list:
      if value == i:
       l.append(value)
  cnt=dict(Counter(l))
  for key, value in sorted(cnt.items(), key=lambda item: (item[0], item[1])):
   print ("Frequency of '"'%s'"': %s" %(key, value))  
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
 