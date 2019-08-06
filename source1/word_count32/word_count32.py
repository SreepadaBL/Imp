from docx import Document
import sys
import os
import win32com.client
import re
from datetime import datetime
from win32com.client import constants
iter=sys.argv[1]
iter1=sys.argv[2]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 34: Checking more than n words in a sentence (n – user input).")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
l=[]
if iter.endswith('.doc'):
 word1 = win32com.client.Dispatch("Word.Application")
 wb = word1.Documents.Open(iter)
 try:
  m=int(iter1)
  if m>0:
   for i in wb.Paragraphs:
    j=i.Range.Text.encode('ascii','ignore').decode()
    k=len(j.split())
    if k>m:
     print("\n")
     print("Text:",j)
     print ("Number of words:",k)
     print("Page number:",i.Range.Information(constants.wdActiveEndAdjustedPageNumber))
     print("Line On Page:",i.Range.Information(constants.wdFirstCharacterLineNumber))
  else:
      print ("Enter an Integer greater than zero!!!")
 except ValueError:
	 print('\nYou did not enter a valid integer')
	 sys.exit(0)
 wb.Quit()  
elif iter.endswith('.docx'):
 c=Document(iter)
 try:
  m=int(iter1)
  if m>0:
   for i in c.paragraphs:
    st=i.style.name
    if re.search("Heading",st):
     l.append(i.text.encode('ascii','ignore').decode())
    #print(l)
    i=i.text.encode('ascii','ignore').decode()
    k=len(i.split())
    if k>m:
     print("\n")
     print("Text:",i)
     print ("Number of words:",k)
     print("Heading Section:",l[-1])
  else:
      print ("Enter an integer greater than zero!!!")
 except ValueError:
	 print('\nYou did not enter a valid integer')
	 sys.exit(0)
else:
	print("invalid")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")