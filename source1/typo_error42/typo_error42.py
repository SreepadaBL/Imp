import win32com.client as win32
import os 
from datetime import datetime
import string
import re
from docx import Document
from win32com.client import constants
##Read the path 
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 36: Checking document for Typo error verification on Run.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
res=[]
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 a=word1.WordBasic.ToolsSpelling
 #print(a)
 #print("Grammatical Errors:",doc.GrammaticalErrors.Count)
 #print("Spelling Errors:",doc.SpellingErrors.Count)
 #for err in doc.SpellingErrors:
 # print(err.Text)
 print("Grammar and Spelling Check Verified in the document")
print(datetime.now()-start,"HH:MM:SS")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
word1.Quit()  