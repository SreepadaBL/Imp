import win32com.client as win32
import os 
import string
import re
from datetime import datetime
from win32com.client import constants
from docx import Document
##Read the path 
import sys
from datetime import datetime                                                                                    
iter=sys.argv[1]
iter2=sys.argv[2]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 24: Document Font Type Check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
count=0
l=[]
ll=[]
lll=[]
if iter.endswith('.doc'):# or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 par=sheet_1.Paragraphs
 
 for para in par:
  a=para.Range.Font.Name
  #print(a.encode('ascii','ignore').decode('unicode_escape'))
  if str(a)!=iter2 and len(str(para).encode('ascii','ignore').decode())>0 and str(para.encode('ascii','ignore').decode()).strip()!='':
   if str(para).encode('ascii','ignore').decode()!='\r\x07':
    #l.append(str(para).encode('ascii','ignore').decode('unicode_escape'))
    print("String:",str(para).encode('ascii','ignore').decode('unicode_escape'))
    print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
    print("\n")
    count=count+1
 if count>0:
  print("Status:Fail")
 else:
  print("Status:Pass")
 sheet_1.Close()
 word1.Quit()  
elif iter.endswith('.docx'):
 doc=Document(iter)
 for p in doc.paragraphs:
  k=p.text.encode('ascii','ignore').decode()
  n=p.style.font.name
  m=p.style.name
  #print(n)
  if re.search("Heading [0-9]",m):
   l.append(p.text.encode('ascii','ignore').decode())
  if n!=iter2 and k.strip()!='' and len(k)>0:
   print("Text:",k)
   print("Font type:","\t",n)
   count=count+1
   if l!=[]:
    print("Heading:","\t",l[-1])
    print("\n")
   else:
    pass
 if count>0:
  print("Status:Fail")
 else:
  print("Status:Pass")
else:
 print("Enter Valid Path!!!!")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
