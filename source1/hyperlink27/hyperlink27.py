import win32com.client as win32
import re
import sys
import os
from win32com.client import constants
from datetime import datetime
count=0
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("-----------------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 13: Hyperlink check for TOC, list of figures and Tables.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("-----------------------------------------------------------------------------------------------------------------")
print("\n")
cnt2=1
flag=0
flag1=0
flag2=0
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 cnt=sheet_1.TablesOfContents.Count
 cnt1=sheet_1.TablesOfFigures.Count
 if cnt==1:
  if sheet_1.TablesOfContents(cnt).UseHyperlinks == True:
   flag=1
  else:
   print("For Table of Contents hyperlink is not used.")
 if cnt2<=cnt1:
  tf=sheet_1.TablesOfFigures(cnt2).Caption
  if tf=="Table":
   tf1=sheet_1.TablesOfFigures(cnt2).UseHyperlinks
   if tf1==True:
    flag1=1
   else:
    print("For List of Tables hyperlink is not used.")
  if tf=="Figure":
   tf1=sheet_1.TablesOfFigures(cnt2).UseHyperlinks
   if tf1==True:
    flag1=1
   else:
    print("For List of Figures hyperlink is not used.")
  cnt2=cnt2+1
 if flag==1 and flag1==1:# and flag2==1:
  print("For TOC,List of Tables and Figures ,Hyperlinks are used")
  print("Status:Pass")
 else:
  print("Status:Fail")
sheet_1.Close()
word1.Quit()
