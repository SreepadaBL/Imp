import win32com.client as win32
import os 
import re
import pythoncom
import sys
from datetime import datetime                                                                                 
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 20: Availability of definition corresponding to acronym in Acronyms and Definition table in Document.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
if iter.endswith('.doc') or iter.endswith('.docx'): 
 word1 = win32.gencache.EnsureDispatch("Word.Application")
 word1.Visible = False
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 list_1=['doc id','version','revision date']
 #print("Yes")

 def get_table_count():
  return sheet_1.Tables.Count
 
 def count_table_rows(table):
  return table.Rows.Count
 
 def count_table_columns(table):
  return table.Columns.Count
 
 def get_headers():
  headers = sheet_1.Sections(1).Headers(1)
  shape_count = headers.Shapes.Count
  for shape_num in range(1, shape_count + 1):
      t_range = headers.Shapes(shape_num).TextFrame.TextRange
      text = t_range.Text
      page_num = t_range.Information(3)  # 3 == wdActiveEndPageNumber
      yield text, page_num
 
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
 
 def __del__():
  word1.Quit()
 #try:
 res=[]
 res1=[]
 
        #path = str(input())
        #count=0
        #open_doc = os.path.abspath(path)
 for table_num, table_text in enumerate(get_all_table_text()):
     #print("\n-------------- Table %s ----------------" % (table_num + 1))
     for row_data in table_text:
         b=", ".join(row_data)
         b=str(b).encode("utf-8")
         #print(b)
         k=b"Acronyms"
         l=b"Definition"
         if k in b: 
             #print(table_text)
             k=table_text[0]
             #print(k)
             m=k.index('Acronyms')
             #print(m)
             for i in table_text:
              aa=i[m]
              if re.sub(r'\s+','',aa):
               res.append(aa)
              bb=i[m+1]              
              if re.sub(r'\s+','',bb):
               res1.append(bb)				
             #print(res1)
              
 #print("Acronyms",res)
 #print("definition",res1) 
 if (len(res)!=len(res1) or (res==[]) or (res1==[])):
  print("Definitions are not updated corresponding to its Acronyms")
  print("Status:Fail")
 else:
  print("Definitions are updated corresponding to its Acronyms")
  print("Status:Pass")
 #except:
 # pass

end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")    
sheet_1.Close()
word1.Quit()