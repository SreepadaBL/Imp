import win32com.client as win32
import os 
import re
import pythoncom
import sys
from collections import Counter
from datetime import datetime                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 16: Revision History Table Blank Field and Date Consistency Check.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'): 
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = False
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 para=sheet_1.Paragraphs
## count the number of tables present in document
 def get_table_count():
  return sheet_1.Tables.Count
## count the number of rows of table present in Document
 def count_table_rows(table):
  return table.Rows.Count

## count the number of columns of table present in Document
 def count_table_columns(table):
  return table.Columns.Count

 ##Reading header content
 def get_headers():
  headers = sheet_1.Sections(1).Headers(1)
  shape_count = headers.Shapes.Count
  for shape_num in range(1, shape_count + 1):
      t_range = headers.Shapes(shape_num).TextFrame.TextRange
      text = t_range.Text
      page_num = t_range.Information(3)  # 3 == wdActiveEndPageNumber
      yield text, page_num
 
 ##Reading content of a table
 def get_table_text(table):
  col_count = count_table_columns(table)
  row_count = count_table_rows(table)
 
  for row in range(1, row_count + 1):
      row_data = []
      for col in range(1, col_count + 1):
          try:
              row_data.append(table.Cell(Row=row,Column=col).Range.Text.strip(chr(7) + chr(13)))
              
          except pythoncom.com_error as error:
              row_data.append("")
 
      yield row_data
 
 ##Reading content of all tables
 def get_all_table_text():
  for table in get_tables():
      table_data = []
      for row_data in get_table_text(table):
          #for col_data in .get_table_text(table):
              #table_data1.append(col_data)
              table_data.append(row_data)
      yield table_data
      #yield table_data1
 
 def get_tables():
  for table in sheet_1.Tables:
      yield table
 
 ##Closing the word Document
 def __del__():
  word1.Quit()
 res=[]
 res1=[]
 final=[]
 final1=[]
 app=[]
 count=0
 ##Read content of all tables present in document
 for table_num, table_text in enumerate(get_all_table_text()):
     #print("\n-------------- Table %s ----------------" % (table_num + 1))
     for row_data in table_text:
         b=", ".join(row_data)    ##concatenate list items to form string and encode it to byte string
         b=str(b).encode("ascii",'ignore').decode().lower()
         #print(b)
         if re.search("^revision",b): 
             k=table_text[0]      ##Accessing first row of a table
             #print(k)
             p=len(k)
             #print(p)
             r = re.compile("^revision",re.IGNORECASE)     
             newlist = list(filter(r.match,k))
             m=k.index(newlist[0])   
             #print(m)
             ppp=m+1
             for i in table_text:
              for j in range(0,p):
               pp=i[m+j].encode('ascii','ignore').decode()
               #print(pp)
               if len(pp)==0:
                print("Blank in column:",j+1)
         
              
               #res.append(pp)
                count=count+1
              if i[ppp]!='date':
               p4=i[ppp].encode('ascii','ignore').decode()
               app.append(p4)
              app1=[] 
              try:				
               for date1 in app[1:]:
                if datetime.strptime(date1, "%Y-%m-%d")==True:
                 pass
              except:
               print("Date Format is incorrect:", date1,"for revision no",i[0])
               #app1.append(p4)
               count=count+1
              #if i[ppp]!='date':
              # p4=i[ppp].encode('ascii','ignore').decode()
              # app.append(p4)				
 #app1=[]				
 #try:				
 # for i in app[1:]:
 #  if i== datetime.strptime(i, "%Y-%m-%d"):
 #   raise ValueError
 #except ValueError:
 # print("Date Format is incorrect:", p4)
 # app1.append(p4)
 #print(app)
 #if res!=[] and app1!=[]:
 if count>=1:
  print("Status:Fail")
 else:
  print("Revision History Table has no blank fields.\nDate Format used is consistent.")
  print("Status:Pass")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
sheet_1.Close()
word1.Quit()    