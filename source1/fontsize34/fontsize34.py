import win32com.client as win32
import os 
import re
from datetime import datetime
from collections import Counter
import string
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx import Document
from win32com.client import constants
import sys
count=0
##Read the path 

seen=[]
seen1=[]
seen2=[]
l=[]
ll=[]
lll=[]
list1=[]
##Open the Document

count=0
l=[]
iter=sys.argv[1]
iter2=sys.argv[2]
iter2=round(float(iter2))
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 25: Document Font Size Check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc'):# or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 par=sheet_1.Paragraphs
 
 #print(list1) 
 try:
  for para in par:
   m=str(para)
   m=m.rstrip('\r\x07')
   m=m.rstrip('\t\r')
   t=str.maketrans({key: None for key in string.punctuation})
   k=m.translate(t)
   j=k.encode("utf-8")
   p=j.rstrip(b'\r\x07')
   if len(p)!=0:
    a=para.Range.Font.Size
    
    if a!=int(iter2) and a!=26 and str(para.Range.Text.encode('ascii','ignore').decode()).strip()!='':
     #l.append(a)
     #ll.append(para.Range.Text.encode('ascii','ignore').decode())
     print("Text:",para.Range.Text.encode('ascii','ignore').decode())
     print("Font-size:",a)
     print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
     print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
     print("-------------------------------------------------------------------------------------------------------------------")
     count=count+1
    #if a==26:
    # seen1.append(str(para))
    ##print(ll)
  #print("Font with size 26:",seen1)
  
  if count==0:
   print("Status:Pass")
  else:
   print("Status:Fail")  
 except:
  pass
 sheet_1.Close()
 word1.Quit()
#print(doc.size)
elif iter.endswith('.docx'):
 doc1 = Document(iter)
 for p in doc1.paragraphs:
  size=p.style.font.size
  k=p.style.name
  if re.search("Heading",k):
   l.append(p.text.encode('ascii','ignore').decode())
  if str(size)!="None":
   size1=p.style.font.size.pt
   if size1!=int(iter2) and size1!=26 and str(p.text.encode('ascii','ignore').decode()).strip()!='':
    print("Text:",p.text.encode('ascii','ignore').decode())
    print("Font Size:",size1)
    count=count+1
    if l!=[]:
     print("Heading:",l[-1])
     print("\n")
    else:
     pass
	 
if count==0:
 print("Status:Pass")
else:
 print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
