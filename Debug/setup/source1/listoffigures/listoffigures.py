import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from win32com.client import constants
count=0
##Read the path 
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 9: List of  Figures update check")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
l=[]
ll=[]
lll=[]
s=[]
ss=[]
sss=[]
if iter.endswith('.doc'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 try:
  cnt=sheet_1.TablesOfFigures.Count
  #print(cnt)
  if cnt>1:
   para=sheet_1.TablesOfFigures(2).Range.Text
   #para1=sheet_1.TablesOfFigures(2).Range.Text
   a=para.splitlines()
   #b=para1.splitlines()
   for item in a:
    item=item.split('\t')
    for i in item:
     if i.isdigit()==False:
      l.append(i.encode('ascii','ignore').decode().lower())
   #print(l)
   
   for para in sheet_1.Paragraphs:  #re.search("[\w\s]?^Table",str(para)) and 
    a=para.Range.Style
    c=para.Range.Text.encode('ascii','ignore').decode()
    if re.search("[\w\s]?Figure[\w\s]?[0-9]",c) and str(a)=='Caption':
     #print(para.Range.Text.encode('ascii','ignore').decode())
     #print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
     b=str(para).strip('\r')
     b=b.strip(' ')
     b=b.strip('\r\x07')
     ll.append(b.encode('ascii','ignore').decode().lower())
     
   #print(ll)
   for i in l:
    if i not in ll:
     print("Figure:",'"',i,'"',"is not updated in List Of Figures.")
     count=count+1
   #q=set(l)-set(ll)
   #print(q)
   if count>0:
    print("Status:Fail")
   else:
    if ll==[] or a==[]:
     print("List Of Figures Not Found in the document.")
    else:
     print("List of Figures are Updated in the Document")
     print("Status:Pass")
  else:
   para=sheet_1.TablesOfFigures(1).Range.Text
   a=para.splitlines()
   for item in a:
    item=item.split('\t')
    for i in item:
     if i.isdigit()==False:
      l.append(i.encode('ascii','ignore').decode().lower())
   #print(l)
   
   for para in sheet_1.Paragraphs:  #re.search("[\w\s]?^Table",str(para)) and 
    a=para.Range.Style
    c=para.Range.Text.encode('ascii','ignore').decode()
    if re.search("[\w\s]?Figure[\w\s]?[0-9]",c) and str(a)=='Caption':
     #print(para.Range.Text.encode('ascii','ignore').decode())
     #print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
     b=str(para).strip('\r')
     b=b.strip(' ')
     b=b.strip('\r\x07')
     ll.append(b.encode('ascii','ignore').decode().lower())
     
   #print(ll)
   for i in l:
    if i not in ll:
     print("Figure:",'"',i,'"',"is not updated in List Of Figures.")
     count=count+1
   #q=set(l)-set(ll)
   #print(q)
   if count>0:
    print("Status:Fail")
   else:
    if ll==[] or a==[]:
     print("List Of Figures Not Found in the document.")
    else:
     print("List of Figures are Updated in the Document")
     print("Status:Pass")
 except:
  print("List Of Figures Not Found in the document.")
 sheet_1.Close()
 word1.Quit()
elif iter.endswith('.docx'):
 doc=Document(iter)
 try:
  body_elements = doc._body._body
  rs = body_elements.xpath('.//w:r')
  table_of_content = [r.text.encode('ascii','ignore').decode() for r in rs if r.style == 'Hyperlink']
  #print(table_of_content)
  for para in doc.paragraphs:
   a=para.style.name
   b=a.encode('ascii','ignore').decode()
   c=para.text.encode('ascii','ignore').decode()
   #print(b)
   if re.search("[\w\s]?Figure[\w\s]?[0-9]",c) and b=='Caption':
    d=c.strip(' ')
    d=d.strip('\r')
    d=d.strip('\n')
    l.append(d.strip(' '))
  #print(l)
  for item in l:
   if item not in table_of_content and item!='':
    print("Figure:",'"',item,'"',"is not updated in List Of Figures.")
    count=count+1
  if count>0:
   print("Status:Fail")
  else:
   if l==[] or table_of_content==[]:
    print("List Of Figures Not Found in the document.")
   else:
    print("List of Figures are Updated in the Document")
    print("Status:Pass")
 except:
  print("List Of Figures Not Found in the document.")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
   