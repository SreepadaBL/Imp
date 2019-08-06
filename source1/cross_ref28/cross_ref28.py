import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx import Document
from win32com.client import constants
import sys
app=[]
iter=sys.argv[1]
start=datetime.now()
res=[]
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 39: Invalid Cross reference ‘Error! Reference source not found’ check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Paragraphs
 for p in para:
  p=p.Range.Text.encode('ascii','ignore').decode()
  if re.search('Error! Reference source not found',p):
   print(p)
   print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
   res.append(p)
 if res==[]:
  print("Status:Pass")
 else:
  print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit() 