#importing packages
import io
from docxtpl import DocxTemplate
import pandas as pd #pandas used to analyse dataframe in spreadsheet
import datetime
import os
from PyPDF2 import PdfFileMerger
from win32com import client
import tkinter as tk
from tkinter import filedialog
import docx


class result_template:
        def __init__(self):
                self.contents = None

        def select_file(self):
                root = tk.Tk()
                root.geometry("550x550")
                root.withdraw()
                root_dir = os.path.dirname(os.path.abspath(__file__))
                file_path = filedialog.askopenfilename(initialdir = root_dir+'/template/', title = "Select a DOCX file", filetypes = (("DOCX files", "*.docx"), ("all files", "*.*")))
                #print(file_path)
                doc = os.path.basename(file_path)
                return doc        
        
        def create_template(self,doc):
                #using client to start MS word application
                word_app = client.Dispatch("Word.Application")
                root_dir = os.path.dirname(os.path.abspath(__file__))
                parent_dir = os.path.dirname(root_dir)

                
                file = root_dir+'/Template/'+doc

                #getting csv as input
                
                csv_file = filedialog.askopenfilename(initialdir = parent_dir + '/xlsx', title = "Select file", filetypes = (("Excel files", "*.csv"), ("all files", "*.*")))

                #reading file from csv format
                #data_frame = pd.read_csv('xlsx/Internals_1_sem_4.csv')
                data_frame = pd.read_csv(csv_file)
                
                
                date_today = datetime.datetime.now().strftime('%d-%m-%y')
                

                # taking heading of a column in as r_index , r_index is used as header or as pointer to each record
                for r_index,row in data_frame.iterrows():
                        S_roll = row['rollno']

                        # For date:
                        data_frame['date']=date_today

                        tpl = DocxTemplate(file)
                        df_doct = data_frame.to_dict()
                        x = data_frame.to_dict(orient='records')
                        content = x
                        


                        #render for transferring content to template by individual report 
                        tpl.render(content[r_index])
                        tpl.save('report\Docx\\'+S_roll+".docx")

                        # time.sleep(1) --> Slow the creation of docs

                        #redirecting the location to project folder so it can used dynamically to store(when used as exe file)
                        root_dir1 = os.path.dirname(os.path.abspath(__file__))


                        #Converting the docx files to pdf for easy merging
                        doc= word_app.Documents.Open(root_dir1+'\\report\Docx\\'+S_roll+'.docx')
                        doc.SaveAs(root_dir1+'\\report\Pdf\\'+S_roll+'.pdf',FileFormat =17)

                print("Exported Successfully!..")
                #close MS Word Application
                word_app.Quit()
                

        def pdf_merger(self):
                #creating merger(like object/instance)
                merger = PdfFileMerger()

                root_dir = os.path.dirname(os.path.abspath(__file__))
                folder_path = root_dir + '/report/Pdf' 
                list_of_files = os.listdir(folder_path)
                for file_item in list_of_files:
                        if '.pdf' in file_item:
                                # print(file_item) --> checking for file availability

                                #using append each file is appended inside single file
                                merger.append(folder_path+'\\'+file_item)

                #Naming the merged file and closing the method
                merger.write("result.pdf")
                merger.close()
                print("Merged Successfully!..")
                
        #def closeEvent(self):
         #   sys.exit(app.exec_())


if __name__ == "__main__":
        
        import sys
        
        app= DocxTemplate(sys.argv)
        run= result_template()
        run.create_template(doc=run.select_file())
        run.pdf_merger()
        sys.exit(0)