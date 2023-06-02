from openpyxl import Workbook
import tabula
import csv
import os
import shutil

#site_url = 'https://larsentoubro-my.sharepoint.com/personal/c50082_lthe_in/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fc50082%5Flthe%5Fin%2FDocuments%2Fpdf%5Fto%5Fexcel&view=0'
#folder_name = 'NewFolder'
# SharePoint credentials (if required)
#username = 'your-username'
#password = 'your-password'
#folder_endpoint = f"{site_url}/_api/web/folders"

file_list = []
files = [f for f in os.listdir('.') if os.path.isfile(f)]
pdf_files = list(filter(lambda f: f.lower().endswith(('.pdf')), files))

for pdf_file in pdf_files:
    pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
    file_list.append(pdf_name)
output_folder = 'C:\\Users\\Admin\\Desktop\\project_1\\Test_final\\outout'
os.mkdir(output_folder)
print(file_list)
for s in file_list: 
  wb = Workbook()
  ws = wb.active
  print(s)
  pdf_path1 = 'C:\\Users\\Admin\\Desktop\\project_1\\Test_final\\'+s+'.pdf'
  tabula.convert_into(pdf_path1, s+ ".csv", output_format="csv", pages='all')
  with open(s+'.csv', 'r') as f:
    reader = csv.reader(f)
    for row in reader:
        ws.append(row)
    wb.save(s+'.xlsx')
  os.remove(s+'.csv')
  shutil.move(pdf_path1, os.path.join(output_folder, s+'.pdf'))
 



 




