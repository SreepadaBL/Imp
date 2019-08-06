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
print("CheckList Rule - 7: Table of Contents update check.")
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
  para=sheet_1.TablesOfContents(1).Range.Text
  a=para.splitlines()
  for item in a:
   item=item.split('\t')
   for i in item:
    l.append(i.encode('ascii','ignore').decode().lower())
  #print(l)
  for para in sheet_1.Paragraphs:  #re.search("[\w\s]?^Table",str(para)) and 
   a=para.Range.Style
   if re.search("[\w\s]?Heading[\w\s]?[0-9]",str(a)):
    #print(para.Range.Text.encode('ascii','ignore').decode())
    #print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    b=str(para).strip('\r')
    b=b.strip(' ')
    b=b.strip('\r\x07')
    ll.append(b.encode('ascii','ignore').decode().lower())
    
  #print(ll)
  for i in ll:
   if i not in l and i!='':
    print("Heading:",'"',i,'"',"is not updated in Table of Contents.")
    print("\n")
    count=count+1
  if count>0:
   print("Status:Fail")
  else:
   if l==[] or a==[]:
    print("Table Of Contents Not Found in the document.")
   else:
    print("Table of Contents are Updated in the Document")
    print("Status:Pass")
 except:
  print("Table Of Contents Not Found in the document.")
 sheet_1.Close()
 word1.Quit()
  #l.append(d[:2])
  #s.append(d[0])
 #for para in sheet_1.Paragraphs:
 # k=para.Range.Text.encode('ascii','ignore').decode()
 # p=k.strip().split('\t')
 # ll.append(p[:2])
 #m=[i for i in ll if i in l]
 #print(m)
 #print(l)
 #dups = {tuple(x) for x in l if l.count(x)>1}
 #dups1 = {x for x in s if s.count(x)>1}
 #print(dups1)
 #z = [tuple(y) for y in m]
 #x = [tuple(y) for y in l]
 #res=set(z)-set(x)
 #if res == set():
 # print("\nAll the headings used in TOC are used in Document\n")
 #else:
 # print(res,"heading of TOC is not used in Document\n")
 #if dups!=set():
 # print("\nDuplicates in TOC\n",dups)
 #if dups1!=set(): 
 # print(dups1,"TOC Number is repeated")
 #if dups==set() and dups1==set(): 
 # print("\nNo Duplicates found in TOC\n")
elif iter.endswith('.docx'):
 doc=Document(iter)
 try:
  body_elements = doc._body._body
  rs = body_elements.xpath('.//w:r')
  table_of_content = [r.text.encode('ascii','ignore').decode() for r in rs if r.style == 'Hyperlink']
  #print(table_of_content)
  for para in doc.paragraphs:
   a=para.style.name
   if re.search("Heading [0-9]",a):
    b=para.text.encode('ascii','ignore').decode()
    b=b.strip(' ')
    b=b.strip('\n')
    b=b.strip('\r')
    b=b.strip('\r\x07')
    b=b.strip('\x0c')
    b=b.strip('\x0b')
    b=b.strip('\x0a')
    b=b.rstrip(' ')
    l.append(b)
  
  #print(l)
  for item in l:
   if item not in table_of_content and item!='':
    print("Heading:",'"',item,'"',"is not updated in Table of Contents.")
    count=count+1
  if count>0:
   print("Status:Fail")
  else:
   if l==[] or table_of_content==[]:
    print("Table Of Contents Not Found in the document.")
   else:
    print("Table of Contents are Updated in the Document")
    print("Status:Pass")
 except:
  print("Table Of Contents Not Found in the document.")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")