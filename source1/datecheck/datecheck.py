import win32com.client as win32
import os 
import re
import pythoncom
import sys
from datetime import datetime          
try:
 iter=sys.argv[1]
 start=datetime.now()
 print("--------------------------------------------------------------------------------------------------------")
 print("Document Name:", iter)
 print("CheckList Rule - 16: 15: Document Version and Date updated properly at Header, Revision History.")
 print("Document Review Start Time:", start,"HH:MM:SS")
 print("--------------------------------------------------------------------------------------------------------")
 print("\n")
 if iter.endswith('.doc') or iter.endswith('.docx'):
  word1 = win32.Dispatch("Word.Application")
  word1.Visible = True
  p = os.path.abspath(iter)
  word1.Documents.Open(p)
  sheet_1 = word1.ActiveDocument
  list_1=['doc id','version','revision date']
  #print("Yes")
  try:
  ## Read table present in header and count no of rows and columns
   HeaderTable=sheet_1.Sections(1).Headers(1).Range.Tables(1)
   HeaderTable_row = sheet_1.Sections(1).Headers(1).Range.Tables(1).Rows.Count
   HeaderTable_column = sheet_1.Sections(1).Headers(1).Range.Tables(1).Columns.Count
   #print (HeaderTable_row)
   #print (HeaderTable_column)
   found =0
   found1=0
   ##list to store the Version,Doc id,Revision Date keywords
   found_array =[]
   ##list to store the corresponding values of Version,Doc id,Revision Date.
   found_array1=[]
   for i in range(1,HeaderTable_row+1):
    for j in range(1,HeaderTable_column+1):
     try:
      #print (i,j)
      a=HeaderTable.Cell(i,j).Range.Text.lower()
      a1=HeaderTable.Cell(i,j).Range.Font.Size
      #print(a)
      #if a1!=10:
      # print(a1,"\n Set the font size to 10")
      for wd in list_1:
       if wd in a:
        #print(word,"keyword present at",i,"row",j,"Column")
        found = found +1
        found_array.append(wd)
		
		## Read adjacent next row of word
		
        b=HeaderTable.Cell(i,j+1).Range.Text
        #print("The",word,"is",b)
        found1=found1+1
        found_array1.append(b)
     except:
      pass
   #print(found_array)
   #print(found_array1)
   #print(found1)
   result=[]
   for item in found_array1:
    item=item.replace("\r\x07","")
    result.append(item)
   #print(result)   
   ## finding missing attributes
   a=set(list_1)-set(found_array)
    #c=set()|set(found_array1)
    #print(c) 
   #if len(a)!=0 or '\r\x07' in found_array1:
   # k=', '.join(a)
   # print(k,"is not present")
   # print("Fail")
   #else:
   # print("Pass")
  
  except:
   sections = sheet_1.Sections
   for section in sections:
    headersCollection = section.Headers
    for header in headersCollection:
        #print (header)
        header=str(header)
        #header1=header.strip()
        header1=header.split('\n')
        #header1=header.split('')
        print(header1)
  
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
   
          #path = str(input())
          #count=0
          #open_doc = os.path.abspath(path)
		  
   ##Read content of all tables present in document
   for table_num, table_text in enumerate(get_all_table_text()):
       #print("\n-------------- Table %s ----------------" % (table_num + 1))
       for row_data in table_text:
           b=", ".join(row_data)       ##concatenate list items to form string and encode it to byte string
           b=str(b).encode("utf-8")
           #print(b)
           k=b"Author"
           #l=b"Author"
           if k in b: 
               #print(table_text)
               k=table_text[0]       ##Accessing first row of a table
               #print(k)
               r = re.compile("^revision",re.IGNORECASE)
               newlist = list(filter(r.match,k))  # Note 1
               #print(newlist)
               m=k.index(newlist[0])
               #print(m)			   ##find index of keyword 'Revision No'
               lm=k.index('Date')        ##find index of keyword 'Date'
               #print(m)
               #res=[]
               for i in table_text:
                #print(i[m])          ##Print 'Revision No' Column
                aa=i[m]
                res.append(aa)       ## store column content in res list
               #res1=[]
               for j in table_text:
                #print(j[lm])		##Print 'Date' Column
                bb=j[lm]
                res1.append(bb)     ## store column content in res1 list
           else:
            break    	
           #print("Pass")
            
   
  except:
   pass
  st=res[-1]
  print("Latest Revision Number in Revision History Table:",st)
  st1=res1[-1]
  print("Revision Date in Revision History Table:",st1)
  print("Revision Number in Header:",result[0])
  print("Revision Date in Header:",result[-1])
  if ((st in result) and (st1 in result)) :
   print("Revision date and Revision Number in Header and Revision History Table are Updated")
   print("Status:Pass")
  else:
   print("Note:Revision date and Revision Number in Header and Revision History Table are not Updated")
   print("Status:Fail")
 else:
  print("Enter the correct path")
except:
 print("Fail")
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")
sheet_1.Close()
word1.Quit()
