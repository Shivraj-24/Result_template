#importing packages
from docxtpl import DocxTemplate
import pandas as pd #pandas used to analyse dataframe in spreadsheet
import datetime
import os
from win32com import client

#using client to start MS word application
word_app = client.Dispatch("Word.Application")


#reading file from csv format
data_frame = pd.read_csv('Internals_1_sem_4.csv')
date_today = datetime.datetime.now().strftime('%d-%m-%y')


# taking heading of a column in as r_index , r_index is used as header or as pointer to each record
for r_index,row in data_frame.iterrows():
        S_roll = row['rollno']
        # For date:
        data_frame['date']=date_today

        tpl = DocxTemplate('Template\ReportTemplate1.docx')        
        df_doct = data_frame.to_dict()
        x = data_frame.to_dict(orient='records')
        content = x

        #render for transferring content to template by individual report 
        tpl.render(content[r_index])
        tpl.save('report\Docx\\'+S_roll+".docx")
        # time.sleep(1) --> Slow the creation of docs

        #redirecting the location to project folder so it can used dynamically to store(when used as exe file)
        root_dir = os.path.dirname(os.path.abspath(__file__))

        #Converting the docx files to pdf for easy merging
        doc= word_app.Documents.Open(root_dir+'\\report\Docx\\'+S_roll+'.docx')
        doc.SaveAs(root_dir+'\\report\Pdf\\'+S_roll+'.pdf',FileFormat =17)

print("Exported Successfully!..")
#close MS Word Application
word_app.Quit()