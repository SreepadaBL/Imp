import win32com.client as win32
import os 
from datetime import datetime
import string
from win32com.client import constants
from docx import Document
import re
count=0
l=[]
##Read the path 
import sys                                                                                  
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 27: Check For Italicized Font Text.(Default:Normal)")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc'):# or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 par=sheet_1.Paragraphs
 try:
  for para in par:
   a=para.Range.Italic
   if a==-1 and len(str(para).encode('ascii','ignore').decode())>0 and str(para.encode('ascii','ignore').decode()).strip()!='':
    print("Italicized Text:",str(para).encode('ascii','ignore').decode())
    print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
    print("\n")
    #l.append(str(para).encode('ascii','ignore').decode())
    count=count+1
  if count>0:
   print("Status:Fail")
  else:
   print("Status:Pass")
 except:
  pass
 sheet_1.Close()
 word1.Quit() 
elif iter.endswith('.docx'):
 doc=Document(iter)
 for p in doc.paragraphs:
  for run in p.runs:
   k=p.style.name
   if re.search("Heading",k):
    l.append(p.text.encode('ascii','ignore').decode())
   if run.italic==True and str(run.text.encode('ascii','ignore').decode()).strip()!='':
    print("Italicized Text:",run.text.encode('ascii','ignore').decode())
    count=count+1
    if l!=[]:
     print("Heading:",l[-1])
     print("\n")
    else:
     pass
 if count>0:
  print("Status:Fail")
 else:
  print("Status:Pass")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")


