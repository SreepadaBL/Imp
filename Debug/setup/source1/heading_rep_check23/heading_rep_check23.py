import win32com.client as win32
import os 
import re
import sys
from collections import Counter
from win32com.client import constants
import sys
from datetime import datetime                                                                                    
iter=sys.argv[1]
start=datetime.now()
count=0
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 11: Heading Row Repetition check for tables.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 try:
  for table in sheet_1.Tables:
   if table.Rows.HeadingFormat==False:  
    print("Table in Page ", table.Range.Information(constants.wdActiveEndAdjustedPageNumber), "on line", table.Range.Information(constants.wdFirstCharacterLineNumber), "has no heading row Repitions\n")
    count=count+1
 except:
  pass
if count>=1:
 print("Status:Pass")
else:
 print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
word1.Quit()  