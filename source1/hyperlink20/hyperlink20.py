import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from win32com.client import constants
from docx.opc.constants import RELATIONSHIP_TYPE as RT
count=0
##Read the path 
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("-----------------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 38: Hyperlinks should not be underlined.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("-----------------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Hyperlinks
 for p in para:
  if str(p.Range.Font.Underline)!='9999999':
   print("Hyperlink:: ", p.Range.Text.encode('ascii','ignore').decode())
   print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
   count=count+1
 if count>0:
  print("Status:Fail")
 else:
  print("No Underlined Hyperlinks Found.")
  print("Status:Pass")
 end=datetime.now()
 print("\nDocument Review End Time:", end)
 print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
 sheet_1.Close()
 word1.Quit()  

#elif iter.endswith('.docx'):
# doc=Document(iter)
# rels=doc.part.rels
# for rel in rels:
#  if rels[rel].reltype == RT.HYPERLINK:
#   print(rels[rel]._target)