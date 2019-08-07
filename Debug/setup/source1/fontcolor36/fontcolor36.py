import win32com.client as win32
import os 
import re
from datetime import datetime
from collections import Counter
import string 
from docx import Document
from docx.shared import RGBColor
from win32com.client import constants
count=0
##Read the path 
import sys                                                                                  
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 26: Document Font Color Check.(Default:Automatic/Black).")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
ll=[]
lll=[]
##Open the Document
if iter.endswith('.doc'):## or iter.endswith('.docx'):
 word1 =win32.gencache.EnsureDispatch('Word.Application')
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 par=sheet_1.Paragraphs

 for para in par:
  m=str(para)
  t=str.maketrans({key: None for key in string.punctuation})
  k=m.translate(t)
  j=k.encode('ascii','ignore').decode() 	
  if re.search(r'[a-z0-9]+',j,re.IGNORECASE):
   a=para.Range.Font.Color
   #print(a)
   if a!=-16777216 and a!=0 and a!=9999999:
    print("Text:",j)
    print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
    print("\n")
    count=count+1
 if count>0:
  print("Status:Fail")
  print("Set Font Color to Automatic\Black.")
 else:
  print("Status:Pass")
 sheet_1.Close()
 word1.Quit()  
elif iter.endswith('.docx'):
 l=[]
 doc=Document(iter)
 para=doc.paragraphs
 for p in para:
  f=p.style.name
  if re.search("Heading [0-9]",f.encode('ascii','ignore').decode()):
   ll.append(p.text.encode('ascii','ignore').decode())
  for run in p.runs :
   a=run.font.color.type
   #print(a)
   if str(a)!='None' and a!=101:
    #print("Text:",p.text.encode('ascii','ignore').decode())
    l.append(p.text.encode('ascii','ignore').decode())
    count=count+1
    if ll!=[]:
     #print("Heading:",ll[-1])
     lll.append(ll[-1])
 cnt=dict(zip(l,lll))
 for key, value in sorted(cnt.items(), key=lambda item: (item[0], item[1])):
  print("Text:",key)
  print("Heading:",value)
  print("\n")  
 if count>0:
  print("Status:Fail")
  print("Set Font Color to Automatic\Black.")
 else:
  print("Status:Pass")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
