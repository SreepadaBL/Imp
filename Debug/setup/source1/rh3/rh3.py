import pythoncom
import re
import sys
import win32com.client as win32
import os
import sys
from win32com.client.gencache import EnsureDispatch
from datetime import datetime
import string
from win32com.client import constants
import re
start=datetime.now()
class OpenDoc(object):
    ##Open the Document
    def __init__(self, docx_path):
        import win32com.client as win32
        self.path = docx_path
        self.word1 = win32.Dispatch("Word.Application")
        self.word1.Visible = True
        self.word1.Documents.Open(iter)
        self.sheet_1 = self.word1.ActiveDocument
    
	## count the number of tables present in document
    def get_table_count(self):
        return self.sheet_1.Tables.Count
    
	## count the number of rows of table present in Document
    def count_table_rows(self, table):
        return table.Rows.Count
    
	## count the number of columns of table present in Document
    def count_table_columns(self, table):
        return table.Columns.Count
    
	##Reading header content
    def get_headers(self):
        headers = self.sheet_1.Sections(1).Headers(1)
        shape_count = headers.Shapes.Count
        for shape_num in range(1, shape_count + 1):
            t_range = headers.Shapes(shape_num).TextFrame.TextRange
            text = t_range.Text
            page_num = t_range.Information(3)  # 3 == wdActiveEndPageNumber
            yield text, page_num
    
	##Reading content of a table
    def get_table_text(self, table):
        col_count = self.count_table_columns(table)
        row_count = self.count_table_rows(table)

        for row in range(1, row_count + 1):
            row_data = []
            for col in range(1, col_count + 1):
                try:
                    row_data.append(table.Cell(Row=row,Column=col).Range.Text.strip(chr(7) + chr(13)))
                    
                except pythoncom.com_error as error:
                    row_data.append("")

            yield row_data

	##Reading content of all tables
    def get_all_table_text(self):
        for table in self.get_tables():
            table_data = []
            for row_data in self.get_table_text(table):
                #for col_data in self.get_table_text(table):
                    #table_data1.append(col_data)
                    table_data.append(row_data)
            yield table_data
            #yield table_data1

    def get_tables(self):
        for table in self.sheet_1.Tables:
            yield table
    
	##Closing the word Document
    def __del__(self):
        self.word1.Quit()


if __name__ == "__main__":
        res=[]
	##Read path from user
        count=0
        iter=sys.argv[1]
        start=datetime.now()
        flag=0
        flag1=0
        print("--------------------------------------------------------------------------------------------------------")
        print("Document Name:", iter)
        print("CheckList Rule - 17: Document Check for Approver Name in RH Table Appropriately.")
        print("Document Review Start Time:", start,"HH:MM:SS")
        print("--------------------------------------------------------------------------------------------------------")
        print("\n")
        flag=0
        if str(iter).endswith('.doc') or str(iter).endswith('.docx'):
         open_doc = OpenDoc(iter)
		##Read content of all tables present in document
         try:
          for table_num, table_text in enumerate(open_doc.get_all_table_text()):
           #print("\n-------------- Table %s ----------------" % (table_num + 1))
           for row_data in table_text:
            b=", ".join(row_data)   ##concatenate list items to form string and encode it to byte string
            b=str(b).encode('ascii','ignore').decode().lower()
            #print(b)
            if re.search("approved by",b): 
             k=table_text[0]    ##Accessing first row of a table
             if re.search("^revision",table_text[0][0],re.IGNORECASE):
              flag=1
              r = re.compile("^approved",re.IGNORECASE)
              newlist = list(filter(r.match,k))   
              m=k.index(newlist[0]) ##find index of keyword 'Approved By'
              for i in table_text:
                pp=i[m].encode('ascii','ignore').decode()
                if len(pp)==0 and i[0]!='' and len(i[0])!=0:
                 print("Blank Cell in column:", m+1, "for revision no", i[0])
                 res.append(pp)
         except:
          pass        			  
        else:
         print("Enter the correct path")
        #print(res[1:])
        if flag==0:
         print("Approver Name in Revision History Table not found.")
        if len(res)>0 or flag==0:
         print("Status:Fail")
        else:
         print("Approver Name in Revision History Table are found.")
         print("Status:Pass")	
        		 
   
end=datetime.now()
print("\nDocument Review End Time:", end)
print("\nTime taken For Document Review:", end-start,"HH:MM:SS")  
del()     