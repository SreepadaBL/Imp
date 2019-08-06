import win32com.client as win32
import os 
from datetime import datetime
import string 
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from win32com.client import constants
##Read the path 
import sys                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 28: Document Printability check")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
res=[]
l=[]
ll=[]
count=0
count1=0
count2=0
##Open the Document
if iter.endswith('.doc') or iter.endswith('.docx'):
 word1 = win32.gencache.EnsureDispatch ("Word.Application")
 word1.Visible = False
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1= word1.ActiveDocument
 page_width=sheet_1.PageSetup.PageWidth
 left_margin=sheet_1.PageSetup.LeftMargin
 right_margin=sheet_1.PageSetup.LeftMargin
 tables_cnt=sheet_1.Tables.Count
 shapes_cnt=sheet_1.Shapes.Count
 inlineshapes_cnt=sheet_1.InlineShapes.Count
 #print(page_width)
 #print(left_margin)
 #print(right_margin)
 #print(tables_cnt)
 #print(shapes_cnt)
 #print(inlineshapes_cnt)
 original=page_width-(left_margin+right_margin)
 #print(original)
 if inlineshapes_cnt>=1:
  """Checking InlineShapes of Document"""
  for i in range(1,inlineshapes_cnt+1):
   shape_width=sheet_1.InlineShapes(i).Width
   if shape_width>original:
    #print("Width:",shape_width)
    sheet_1.InlineShapes(i).Select
    a=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
    print("InlineShape:",i,"width crossed the margin in page:",a)
    count=count+1
 if shapes_cnt>=1:
  """Checking Shapes or Pictures of Document"""
  for j in range(1,shapes_cnt+1):
   picture_width1=sheet_1.Shapes(j).Width
   picture_wrap=sheet_1.Shapes(j).WrapFormat.Type
   if picture_wrap==7:
    if picture_width1>original:
     #print("Width:",shape_width)
     sheet_1.Shapes(j).Select
     c=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
     print("Picture:",j,"width crossed the margin in page:",c)
     count1=count1+1
   else:
     picalign=sheet_1.Shapes(j).Left
     if picalign==-999995 or picalign==-999996 or picalign==-999998:
      if picture_width1>original:
       sheet_1.Shapes(j).Select
       c=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
       print("Picture:",j,"width crossed the margin in page:",c)
       count1=count1+1
     elif picalign==-999995 and picalign==-999996 and picalign==-999998 and picalign>=0:
      other=picture_width1+picalign
      sheet_1.Shapes(j).Select
      c=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
      print("Picture:",j,"width crossed the margin in page:",c)
      count1=count1+1
     elif(picalign<-3.6):
      sheet_1.Shapes(j).Select
      c=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
      print("Picture:",j,"width crossed the margin in page:",c)
      count1=count1+1
 
 if tables_cnt>=1:
  width=0
  """Checking Table Width of Document"""
  try:
   for k in range(1,tables_cnt+1):
    table_width=sheet_1.Tables(k).Columns.Count
    for m in range(1,table_width+1):
     cell=sheet_1.Tables(k).Cell(1,m).Range.Text
     pwidth=sheet_1.Tables(k).Cell(1,m).Width
     width=width+pwidth
     alignment=sheet_1.Tables(k).Rows.Alignment
     if alignment==0:
      indent=sheet_1.Tables(k).Rows.LeftIndent
      if indent>=0:
       e=width+indent
       if e>original:
        sheet_1.Tables(k).Select
        d=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
        print("Table:",k,"width crossed the margin in page:",d)
        count2=count2+1
       else:
        sheet_1.Tables(k).Select
        d=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
        print("Table:",k,"width crossed the left margin in page:",d)
        count2=count2+1		
     if alignment==1:
      if width>original:
       sheet_1.Tables(k).Select
       d=sheet_1.ActiveWindow.Selection.Information(constants.wdActiveEndAdjustedPageNumber)
       print("Table:",k,"width crossed the margin in page:",d)
       count2=count2+1
  except:
   pass 
if count==0 and count1==0 and count2==0:
 print("All Shapes and text in Document are within the Paper margin.Document is printable.")
 print("Status:Pass")
else:
 print("Document is not printable.") 
 print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit()  
	 