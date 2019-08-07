from win32com.client import constants, Dispatch
import string, os
import sys
from datetime import datetime                                                                                    
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 5: Display Document Properties for self-verification. (Author's Name, Company and Title).")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 wd = Dispatch('Word.Application')
 wd.Documents.Open(iter)
 sheet_1 = wd.ActiveDocument
 ##store the properties which is needed to be checked 
 myprops = ["Author", "Company","Title"]
 ##Access built in word document properties
 for prop in myprops:
  print (prop,":", sheet_1.BuiltInDocumentProperties(prop))
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
wd.Quit()   