import win32com.client as win32
import os
import sys
from win32com.client.gencache import EnsureDispatch
from datetime import datetime
from win32com.client import constants
import re
ll=[]
list1=[]
count=0
iter=sys.argv[1]
start=datetime.now()
flag=0
flag1=0
print("---------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 2: Check for 'TSIP Confidential' information, page number with prefix 'page', latest document version in Document Footer is performed.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("---------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'):
	word1 = win32.gencache.EnsureDispatch('Word.Application')
	word1.Visible = True
	p = os.path.abspath(iter)
	word1.Documents.Open(p)
	app=[]
	a=word1.ActiveDocument.Sections(1).Footers(win32.constants.wdHeaderFooterPrimary).Range.Text.encode('ascii','ignore').decode()
	m=a.strip().splitlines()
	#print(m)
	for i in m:
		j=i.lower()
		list1.append(j)
		l=j.split('\t')
	
	print("Footer in the document:\n")
	print(','.join(list1))
	#print(l)
	#print(len(l))
	k=["tsip confidential","tsip generally","tsip strictly"]
	for q in k:
		if q in l:
			flag=1
			break
			
		
		
	for i in l:
		if(re.search("^page",str(i))) and (re.search("of",str(i))):
			ll.append(i)
			flag1=1
			break
			
if flag!=1:
	print("TSIP Confidential/TSIP Generally/TSIP Strictly not found")
if flag1!=1:
	print("Page Number should be in the format of Page X of Y")
if flag==0 or ll==[]:
	print("\nStatus:Fail")
else:
	print("\nStatus:Pass")

end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")

word1.Quit()