import win32com.client as win32
import os 
import re
import sys
import time
from datetime import datetime                                                                              
iter=sys.argv[1]
start=datetime.now()
print("--------------------------------------------------------------------------------------------------------")
print("Document Name:", iter)
print("CheckList Rule - 18: Alphabetical order of acronyms in Acronyms and Definition table in Document.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")

try:
 ##Open the Document and read
 if iter.endswith('.doc') or iter.endswith('.docx'): 
  word1 = win32.Dispatch("Word.Application")
  word1.Visible = True
  p = os.path.abspath(iter)
  word1.Documents.Open(p)
  sheet_1 = word1.ActiveDocument
  try:
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
   
   ##Read content of all tables present in document
   for table_num, table_text in enumerate(get_all_table_text()):
       #print("\n-------------- Table %s ----------------" % (table_num + 1))
       for row_data in table_text:
           b=", ".join(row_data)    ##concatenate list items to form string and encode it to byte string
           b=str(b).encode("utf-8")
           #print(b)
           k=b"Acronyms"
           if k in b: 
               #print(table_text)
               k=table_text[0]      ##Accessing first row of a table
               #print(k)
               r = re.compile("^acronym",re.IGNORECASE)      ##find index of keyword 'Acronyms'
               newlist = list(filter(r.match,k))
               m=k.index(newlist[0])   
               #print(m)
               for i in table_text:
                #print(i[m])
                aa=i[m]
                res.append(aa)    ## store column content in res list
    
  except:
   pass
  #print(res)
  for i in res[1:]:
   #print(i)
   m=i[0].lower()
   #print(m)
   final.append(m)
  ##print("Original:",final)
  ##print("Sorted:",sorted(final))
  if((final!=[]) or (sorted(final)!=[]) and (sorted(final)==final)):
   print("Acronyms in 'Acronyms and Definition' Table are in Alphabetical Order")
   print("Status:Pass")
  else:
   print("Acronyms in 'Acronyms and Definition' Table are not in Alphabetical Order")
   print("Status:Fail")
   
 else:
  print("Enter valid path.")
except:
 print("Status:Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit()