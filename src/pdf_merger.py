#importing required packages
import os
from PyPDF2 import PdfFileMerger

#creating merger(like object/instance)
merger = PdfFileMerger()


folder_path = r'C:/Workings/Shivaraj/Projects/Python/Test 1/report/Pdf'
list_of_files = os.listdir(folder_path)
for file_item in list_of_files:
    if '.pdf' in file_item:
        # print(file_item) --> checking for file availability
        
        #using append each file is appended inside single file
        merger.append(folder_path+'\\'+file_item)
        
#Naming the merged file and closing the method
merger.write("result.pdf")
merger.close()