import win32com.client as win32
import os 
import re
import ntpath
from datetime import datetime
from docx import Document
import sys
from datetime import datetime  
from win32com.client import constants                                                                                  
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 35: Unit check (Kbytes/Mbytes or kHz/MHz) .")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
count=0
l=[]
m=['kb','kB','Kb','KB','Mb','MB','mb','mB','kBytes','mBytes']
n=['KHz','khz','Khz','Mhz','mhz','mHz']
##Open the Document
if iter.endswith('.doc'):# or path.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 for p in sheet_1.Paragraphs:
  for j in m:
   if re.search(j + r"\b" ,p.Range.Text.encode('ascii','ignore').decode()):
    print('"',j,'"',"is used instead of KBytes/Mbytes")
    print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
    print("\n")
    count=count+1
  for j in n:
   if re.search(j + r"\b" ,p.Range.Text.encode('ascii','ignore').decode()):
    print('"',j,'"',"is used instead of kHz/MHz")
    print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
    print("\n")
    count=count+1
 if count==0:
  print("Status:Pass")
 else:
  print("Status:Fail")
 sheet_1.Close()
 word1.Quit() 
elif iter.endswith('.docx'):
 doc=Document(iter)
 for p in doc.paragraphs:
  f=p.style.name
  if re.search("Heading [0-9]",f.encode('ascii','ignore').decode()):
   l.append(p.text.encode('ascii','ignore').decode())
  for j in m:
   if re.search(j + r"\b" ,p.text.encode('ascii','ignore').decode()):
    print('"',j,'"',"is used instead of KBytes/Mbytes")
    print("Text:",p.text.encode('ascii','ignore').decode())
    count=count+1
    if l!=[]:
     print("Heading:",l[-1])
     print("\n")
  for j in n:
   if re.search(j + r"\b",p.text.encode('ascii','ignore').decode()):
    print('"',j,'"',"is used instead of kHz/MHz")
    print("Text:",p.text.encode('ascii','ignore').decode())
    count=count+1
    if l!=[]:
     print("Heading:",l[-1])
     print("\n")
 if count==0:
  print("Status:Pass")
 else:
  print("Status:Fail")
  #print("Use Kbytes/Mbytes or kHz/MHz")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS") 