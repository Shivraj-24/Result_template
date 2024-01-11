# Importing package tabula ( for analysing table data's )
import tabula
  
# Read PDF File
df = tabula.read_pdf("C:\Workings\Shivaraj\Projects\Python\Test 1\doc\pdf to excel\Internals_1_sem_4.pdf", pages = 'all')[0]
  
# Convert into Excel File
df.to_csv('Internals_1_sem_4.csv')