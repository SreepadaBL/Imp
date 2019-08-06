import win32com.client
import os 
import re
import pythoncom
import sys
from collections import Counter
from win32com.client import constants
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
app=[]
count=0
l=[]
import sys
from datetime import datetime                                                                                    
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 10: Table name should not be bold.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc'): #or iter.endswith('.docx'):
 word1 = win32com.client.gencache.EnsureDispatch ("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Paragraphs
 for p in para:
  
  k=p.Range.Text.encode('ascii','ignore').decode()
  #print(p.Range.Style,k)
  #if p.Range.Hyperlinks and re.search("^Table",str(p)) and p.Range.Font.Bold:
  # print("Bold Table name::", k)
  # count=count+1
  # print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
  # print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
  if p.Range.Font.Bold and re.search("^Table",str(p)) : 
   if str(p.Range.Style)=='Caption':  # and p.Range.Font.Bold and p.Range.Style=='Caption'
    print("Bold Table name:",k)
    print("Page number:",p.Range.Information(constants.wdActiveEndAdjustedPageNumber))
    print("Line On Page:",p.Range.Information(constants.wdFirstCharacterLineNumber))
    count=count+1
 if count==0:
  print("Status:Pass")
 else:
  print("Status:Fail")
 sheet_1.Close()
 word1.Quit() 
 
elif iter.endswith('.docx'):
 doc = Document(iter)
 for para in doc.paragraphs:
  f=para.style.font.bold
  if f==True and re.search("^Table",para.text.encode('ascii','ignore').decode()):
   print(para.text.encode('ascii','ignore').decode())
   count=count+1
 if count==0:
  print("Status:Pass")
 else:
  print("Above are the Table Names with bold font.")
  print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS") 
