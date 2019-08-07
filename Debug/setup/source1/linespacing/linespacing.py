import win32com.client as win32
import os 
import re
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Length,Pt
import sys
from datetime import datetime                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 31: Line Spacing Consistency Check throughout the document.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
count=0
l=[]
if iter.endswith('.doc'): 
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 try:
  for para in sheet_1.Paragraphs:
   if len(para.Range.Text.encode('ascii','ignore').decode())>0:
    a=para.LineSpacing
    count=count+1
    l.append(a)
  print("Line Spacing Used:",set(l))
 except:
  pass
 if len(set(l))>1:
  print("Line Spacing is inconsistent.")
  print("Status:Fail")
 else:
  print("Line Spacing is Consistent ")
  print("Status:Pass")
 sheet_1.Close()
 word1.Quit()  
elif iter.endswith('.docx'):
 doc = Document(iter)
 for para in doc.paragraphs:
  if len(para.text.encode('ascii','ignore').decode())>0:
   #print(para.text.encode('ascii','ignore').decode(), para.paragraph_format.line_spacing_rule)
   a=para.paragraph_format.line_spacing_rule
   #print(a)
   l.append(a)   
 #print("line Spacing Used:",l)
 print("Line Spacing Used:",set(l))
 if len(set(l))>1:
  print("Line Spacing is inconsistent.")
  print("Status:Fail")
 else:
  print("Line Spacing is Consistent ")
  print("Status:Pass")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")