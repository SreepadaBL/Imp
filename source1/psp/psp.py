import win32com.client as win32
import os 
import win32com.client
import re
import pythoncom
import sys
from datetime import datetime
from win32com.client import constants                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 29: Page setup check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'): 
 word1 = win32com.client.gencache.EnsureDispatch ("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 if sheet_1.PageSetup.Orientation==constants.wdOrientPortrait:
  print("Page Orientation of this document is Potrait")
 else:
  print("Page Orientation of this document is LandScape")
 if sheet_1.PageSetup.PaperSize==constants.wdPaperA4:
  print("PaperSize of this document is A4")
  print("Status:Pass")
 else:
  print("Set PaperSize to A4")
  print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit() 