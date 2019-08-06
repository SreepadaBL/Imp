import win32com.client as win32
import os 
import re
from docx import Document
import os
import sys
from win32com.client.gencache import EnsureDispatch
from datetime import datetime
import string
from win32com.client import constants
import sys
iter=sys.argv[1]
start=datetime.now()
print("---------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 1: Availability of Document ID, Document Version and Date in Header is checked.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("---------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True      ##The document is opened in background and is not visible 
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 list_1=['doc id','version','revision date']
 #print("Yes")
 try:
  ## Read table present in header and count no of rows and columns
  HeaderTable=sheet_1.Sections(1).Headers(1).Range.Tables(1)
  HeaderTable_row = sheet_1.Sections(1).Headers(1).Range.Tables(1).Rows.Count
  HeaderTable_column = sheet_1.Sections(1).Headers(1).Range.Tables(1).Columns.Count
  #print (HeaderTable_row)
  #print (HeaderTable_column)
  found =0
  found1=0
  ##list to store the Version,Doc id,Revision Date keywords
  found_array =[]
  ##list to store the corresponding values of Version,Doc id,Revision Date.
  found_array1=[]
  for i in range(1,HeaderTable_row+1):
   for j in range(1,HeaderTable_column+1):
    try:
     #print (i,j)
     a=HeaderTable.Cell(i,j).Range.Text.lower()
     #print(a)
     if re.search("^[\w\s]?doc[\.]?[\w\s]?id",a) or re.search("^[\w\s]?version",a) or re.search("^[\w\s]?revision date",a):#^[\w\s]+?doc[\.\w\s]+?id
      a=a.strip('\r\x07')
      a=a.strip('\r')
      a=a.strip('\r\x07')
      a=a.strip('\x0c')
      a=a.strip('\x0b')
      a=a.strip('\x0a')
      a=a.rstrip(' ')
      a=a.strip('\n')
      
      #print("The",a.encode('ascii','ignore').decode(),"keyword present",i,"row",j,"column")
      found = found +1
      found_array.append(a)
   ## Read adjacent next row of word 
      b=HeaderTable.Cell(i,j+1).Range.Text
      b=b.strip('\r\x07')
      b=b.strip('\r')
      b=b.strip('\r\x07')
      b=b.strip('\x0c')
      b=b.strip('\x0b')
      b=b.strip('\x0a')
      b=b.rstrip(' ')
      b=b.strip('\n')
      #print(b.encode('utf-8'))
      if b!='':
       print("The",a,"is:","\t",b.encode('ascii','ignore').decode())
       found1=found1+1
       found_array1.append(b)
    except:
     pass
  if len(found_array)==3 and len(found_array1)==3:
   print("Status:Pass")
  else:
   print("Note: Header is not in accordance with QMS.")
   print("Status:Fail")
  
 except:
  print("Header not found in the Document.")
  print("Status:Fail")
else:
 print("Enter the correct path")
 
 #header = section.header
 #for paragraph in header.paragraphs:
 #    print(paragraph.text) 
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit()