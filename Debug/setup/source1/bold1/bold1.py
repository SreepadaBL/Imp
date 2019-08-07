import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from win32com.client import constants


##Read the path 
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("-----------------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 40:Check Whether cross referred text is in bold in Document.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("-----------------------------------------------------------------------------------------------------------------")
print("\n")

count=0
l=[]
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Fields
 for p in para:
  k=p.Code.Text
  #print(str(k))
  #if (k.find("/REF _Ref/")):
  if re.search("/REF _Ref/",k):
   fon=p.Code.Font.Bold
   if fon==0 and str(p.Code.Style)!='Hyperlink' and str(p.Code.Style)!='Caption':# and str(p.Code.Style)!='TOC 1' and str(p.Code.Style)!='TOC 2' and str(p.Code.Style)!='TOC 3' and str(p.Code.Style)!='Table of Figures':
    ft=p.Result.Text.encode('ascii','ignore').decode()
    print("Cross Referred Text not in Bold:",ft)
    print("Page number:",p.Code.Information(constants.wdActiveEndAdjustedPageNumber))
    print("\n")
    count=count+1
if count>0:
 print("Status:Fail")
else:
 print("Status:Pass")
 print("Cross Referred Text are in Bold.")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
word1.Quit() 
 