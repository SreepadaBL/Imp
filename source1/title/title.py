import glob
import win32com.client as win32
import re
from win32com.client import constants, Dispatch
import string, os   
import sys
from datetime import datetime
import string         
from docx import Document                                                                         
iter=sys.argv[1]
start=datetime.now()
l=[]
j=[]
count=0
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 4: Document Name is available at very first line of Title Sheet.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para1=sheet_1.Paragraphs
 title=sheet_1.BuiltInDocumentProperties('Title')
 if title!='':
  print("Title of the Document:",title)
  count=count+1
 elif title=='':
  for p in para1:
   d=p.Range.Text.encode('ascii','ignore').decode()
   d=d.strip('\r')
   d=d.strip('\r\x07')
   d=d.strip('\x0c')
   d=d.strip('\x0b')
   d=d.strip('\x0a')
   d=d.rstrip(' ')
   d=d.strip('\n')
   #if re.search('Software Design Specification',str(d)):
   # print(d)
   # break
   if re.search('User Manual',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Statement of Work|Scope of Work',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Project Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Design Specification',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Requirement Specification|Software Requirements Specification',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Test Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Release Notes',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Performance Management Report ',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Process Performance Baseline',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Development Interface Agreement',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Quality Assurance Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Maintenance Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Low Level Design',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('High Level Design',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Configuration Management Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Plan for',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Performance Management Report',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Solution Engineering Document',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Development Interface Agreement',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Quality Assurance Plan',d,re.IGNORECASE):
    l.append(d)
    break
  print("Title of the Document:",str(l[0]).replace('[','').replace(']',''))
 
 if l==[] and count==0:
  print("Document Has No Title")
  print("Status:Fail")
 else:
  print("Status:Pass")
 end=datetime.now()
 print("\nDocument Review End Time:", end)
 print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
 sheet_1.Close()
 word1.Quit()
elif iter.endswith('.docx'):
 sheet_1=Document(iter)
 a=sheet_1.core_properties
 title1=a.title
 if title1!='':
  print("Title of the Document:",title1)
  count=count+1
 elif title1=='':
  for p in sheet_1.paragraphs:
   d=p.text.encode('ascii','ignore').decode()
   if re.search('User Manual',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Statement of Work|Scope of Work',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Project Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Design Specification',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Requirement Specification|Software Requirements Specification',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Test Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Release Notes',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Performance Management Report ',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Process Performance Baseline',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Development Interface Agreement',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Quality Assurance Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Maintenance Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Low Level Design',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('High Level Design',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Configuration Management Plan',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Plan for',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Organisation Performance Management Report',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Solution Engineering Document',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Development Interface Agreement',d,re.IGNORECASE):
    l.append(d)
    break
   elif re.search('Software Quality Assurance Plan',d,re.IGNORECASE):
    l.append(d)
    break
 
 for i in l:
  print("Title of the Document:","'",i,"'")
 if l==[] and count==0:
  print("Document Has No Title")
  print("Status:Fail")
 else:
  print("Status:Pass")
 end=datetime.now()
 print("\nDocument Review End Time:", end)
 print("\nTime taken For Document Review:", end-start,"HH:MM:SS") 
 