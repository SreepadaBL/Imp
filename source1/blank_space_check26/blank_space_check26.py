import win32com.client as win32
import os 
import re
import pythoncom
import sys
from datetime import datetime
from win32com.client import constants    
from docx import Document                                                                             
iter=sys.argv[1]
start=datetime.now()
count=0
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 30:Blankspace Check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
l=[]
if iter.endswith('.doc') :
 word1 = win32.gencache.EnsureDispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para = sheet_1.Paragraphs
 for p in para:
  k=p.Range.Text.encode('ascii','ignore').decode()
  if re.search('([A-Z]|[a-z]|[0-9])\s\s+',k):
   print("Double Space found in line::" ,k)
   print("Page Number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   print("Line Number:",p.Range.Information(constants.wdFirstCharacterLineNumber))
   count=count+1
  if re.search('^\s+([A-Z]|[a-z]|[0-9])',k) or re.search('^\s\s+',k):
   print("Space found at the beginning of ::  ",k)
   print("Page Number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   print("Line Number:",p.Range.Information(constants.wdFirstCharacterLineNumber))
   count=count+1
 if count>0:
  print("Status:Fail")
 else:
  print("Status:Pass")
 sheet_1.Close()
 word1.Quit() 
elif iter.endswith('.docx'):
 document = Document(iter)
 for para in document.paragraphs:
  s=para.style.name
  if re.search("Heading [0-9]",str(s)):
   l.append(para.text.encode('ascii','ignore').decode())
  k=para.text.encode('ascii','ignore').decode()
  if re.search('([A-Z]|[a-z]|[0-9])\s\s+',k):
   print("Double Space found in line::" ,"\t",k)
   if l!=[]:
    print("Heading Section:","\t",l[-1])
    print("\n")
    count=count+1
   #print("Page Number:",.Information(constants.wdActiveEndAdjustedPageNumber))
   #print("Line Number:",.Information(constants.wdFirstCharacterLineNumber))
  if re.search('^\s+([A-Z]|[a-z]|[0-9])',k) or re.search('^\s\s+',k):
   print("Space found at the beginning of ::  ","\t",k)
   if l!=[]:
    print("Heading Section:","\t",l[-1])
    print("\n")
    count=count+1
   #print("Page Number:",.Information(constants.wdActiveEndAdjustedPageNumber))
   #print("Line Number:",.Information(constants.wdFirstCharacterLineNumber))
 if count>0:
  print("Status:Fail")
 else:
  print("Status:Pass")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")