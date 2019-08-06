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
print("CheckList Rule - 22: Probable Acronyms in Document not defined in Acronym and Definition Table.")
print("Document Review Start Time:", start,"HH:MM:SS")
print("--------------------------------------------------------------------------------------------------------")
print("\n")
app=[]
if iter.endswith('.doc') or iter.endswith('.docx'): 
 word1 = win32.Dispatch("Word.Application")
 word1.Visible = True
 p = os.path.abspath(iter)
 word1.Documents.Open(p)
 sheet_1 = word1.ActiveDocument
 for para in sheet_1.Paragraphs:
  a=para.Range.Text.encode('ascii','ignore').decode()
  pattern = r'(?:[A-Z]\.)+'
  t = re.findall('([A-Z]+)',a)
  for i in t:
   if len(i)>=2:     #discarding single capital letters
    app.append(i)
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
 try:
  res=[]
  res1=[]
  pp={}
  jj=[]
  jjj=[]
  jjjj=[]
  
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
               res.append(i[m])
  print("----------------------------------------------------------------------------------------------------------------------")		
  print("Acronyms Defined:\n",res[1:])
  print("----------------------------------------------------------------------------------------------------------------------")
  res1=res[1:]
  print("Probable Acronyms not defined:\n",set(app)-set(res1))
  pp=set(app)-set(res1)
  #if len(pp)==0:
  # print("\nStatus:Pass")
  #else:
  # print("\nStatus:Fail")
 except:
  pass 

end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS") 
sheet_1.Close()
word1.Quit()   