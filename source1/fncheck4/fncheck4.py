import win32com.client as win32
import os 
import re
import ntpath
from datetime import datetime
import sys
from colorama import init
from termcolor import colored
from colorama import Fore, Back, Style
iter=sys.argv[1]
start=datetime.now()
print("---------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 3: Document ID check in Filename and Header.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("---------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 list_1=['sheet_1 id','version','revision date','Doc ID','Doc id']
 flag=0
 ## Read table present in header and count no of rows and columns
 HeaderTable=sheet_1.Sections(1).Headers(1).Range.Tables(1)
 HeaderTable_row = sheet_1.Sections(1).Headers(1).Range.Tables(1).Rows.Count
 HeaderTable_column = sheet_1.Sections(1).Headers(1).Range.Tables(1).Columns.Count
 found_array1=[]

 #list to store the corresponding values of Version,Doc id,Revision Date.
 for i in range(1,HeaderTable_row+1):
  for j in range(1,HeaderTable_column+1):
   try:
    #print (i,j)
    
    a=HeaderTable.Cell(i,j).Range.Text.lower()
    a1=HeaderTable.Cell(i,j).Range.Font.Size
    #print(a)
   
    if re.match("(doc)[^\w]+(id)",a,re.IGNORECASE):
     #print("DOC ID:",a)
     b=HeaderTable.Cell(i,j+1).Range.Text.lower().encode('ascii','ignore').decode()
     #print("The",word1,"is",b.upper())
     found_array1.append(b.upper())
     #print(found_array1)
     flag=1
     
   except:
    continue
 if flag==1:
  m=ntpath.basename(iter)
  s=m.split('.doc')
  #print(s[0])
  
  for p in found_array1:
   p=p.rstrip('\r\x07')
   #print(p)
   if p==s[0].upper():
    print("Document ID in header:","\t",p)
    #print('\033[31m' + 'some red text')
	
    print("Filename:","\t",s[0].upper())
    print("\nDocument ID in header is same as Document Filename.")
    init()
    print('\033[31m'+"\nStatus:Pass")
    print(Style.RESET_ALL)
   else:
    print("Document ID in header:","\t",p)
    print("Filename:","\t",s[0].upper())
    print("\nDocument ID in header is not same as Document Filename.")
    init()
    print("\nStatus:Fail")
    print(Style.RESET_ALL)
 else:
  print("'Document ID' Keyword is not present")
  init()
  print("Status:Fail")
  print(Style.RESET_ALL)
else:
 print("Enter the correct path")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit()
