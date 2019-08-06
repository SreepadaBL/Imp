import win32com.client
import os 
import re
import sys
from sys import exit
from win32com.client import constants
import docx
from docx.enum.text import WD_BREAK
from datetime import datetime
app=[]
count=0
l=[]
ll=[]
lll=[]
res=[]
fig=[]
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("-----------------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 37:Appendix reference Check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("-----------------------------------------------------------------------------------------------------------------")
print("\n")
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32com.client.gencache.EnsureDispatch ("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para = sheet_1.Hyperlinks
 para1=sheet_1.Paragraphs
 try:
  for p in para:
   m=p.Range.Text.encode('ascii','ignore').decode()
   #print(m)
   if re.search('[\w\s]+?APPENDIX',m) or re.search('^APPENDIX',m):
    #print(m)
    f=m.split('\t')
  for i in f:
   f1=i.strip('\r')
   f1=f1.strip('\r\x07')
   f1=f1.strip('\x0c')
   f1=f1.strip('\x0b')
   f1=f1.strip('\x0a')
   f1=f1.rstrip(' ')
   f1=f1.strip('\n')
   f1=f1.strip('\r\x07')
   app.append(f1.lower())

  if app==[]:
   print("No Appendix Found in the Document Table of Contents.")
   end=datetime.now()
   print("\nDocument Review End Time:", end)
   print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
   sheet_1.Close()
   word1.Quit() 
   exit(0)
 except:
 # #end=datetime.now()
 # #print("\nDocument Review End Time:", end)
 # #print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
 # #doc.Close()
 # #word1.Quit() 
 # #exit(0)
  pass
 for para in para1:	
  a=para.Range.Font.Bold
  b=para.Range.Style
  #print(b)
  c=para.Range.Text.encode('ascii','ignore').decode()
  if str(a) == '-1' and re.search("appendix",c,re.IGNORECASE) and re.search("^Normal",str(b)) : #or re.search("^appendix",c,re.IGNORECASE): #
   #print(c)
   #print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   #print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
   c=c.strip('\r')
   c=c.strip('\r\x07')
   c=c.strip('\x0c')
   c=c.strip('\x0b')
   c=c.strip('\x0a')
   c=c.rstrip(' ')
   c=c.strip('\n')
   l.append(c.lower())
   ll.append(para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   lll.append(para.Range.Information(constants.wdFirstCharacterLineNumber))  
  if str(a) == '-1' and re.search("appendix",c,re.IGNORECASE) and re.search("^Heading",str(b)): 
   #print(c)
   #print("Page number:",para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   #print("Line On Page:",para.Range.Information(constants.wdFirstCharacterLineNumber))
   c=c.strip('\r')
   c=c.strip('\r\x07')
   c=c.strip('\x0c')
   c=c.strip('\x0b')
   c=c.strip('\x0a')
   c=c.rstrip(' ')
   c=c.strip('\n')
   l.append(c.lower())
   ll.append(para.Range.Information(constants.wdActiveEndAdjustedPageNumber))
   lll.append(para.Range.Information(constants.wdFirstCharacterLineNumber))
 
 for i in l:
  if i in app:
   count=count+1
  else:
   res.append(i)
 if res==[]:
  print("Appendix in Table of Contents are referred in Document.")
  print("Status:Pass")
 else:
  print("Appendix Not Referred:\n",res)
  print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
word1.Quit() 
 
   