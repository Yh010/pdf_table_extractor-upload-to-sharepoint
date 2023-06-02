
#https://www.youtube.com/watch?v=dkxVTX5Hs_Q
#https://github.com/iamlu-coding/python-sharepoint-office365-api/blob/main/examples/download_files.py
#https://www.youtube.com/watch?v=w0pBFo9zpiU

from office365_api import SharePoint
import re
import sys, os, io
import pandas as pd
from openpyxl import Workbook
import tabula
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file_creation_information import FileCreationInformation


# SharePoint settings
site_url = 'https://your-sharepoint-site-url'
folder_url = '/sites/your-site/documents/library-folder'  # The URL of the SharePoint library folder to upload to
username = 'your-username'
password = 'your-password'


# Create a SharePoint client context
ctx = ClientContext(site_url).with_credentials(username, password)


# 1 args = SharePoint folder name. May include subfolders YouTube/2022
FOLDER_NAME = sys.argv[1]

# 2 args = SharePoint file name. This is used when only one file is being downloaded
# If all files will be downloaded, then set this value as "None"
FILE_NAME = sys.argv[2]
# 3 args = SharePoint file name pattern
# If no pattern match files are required to be downloaded, then set this value as "None"
FILE_NAME_PATTERN = sys.argv[3]


def modify_file(file_odj):
    data = io.BytesIO(file_odj)
    


#file_name => actual file name of the pdf with extension
#folder_name => where the converted excel needs to be uploaded
#context => actual file data itself that will be pushed
def upload_file(file_name, folder_name, content):
    SharePoint().upload_file(file_name, folder_name, content)


def get_file(file_n, folder):
    file_obj = SharePoint().download_file(file_n, folder)
    

def get_files(folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        get_file(file.name, folder)

def get_files_by_pattern(keyword, folder):
    files_list = SharePoint()._get_files_list(folder)
    for file in files_list:
        if re.search(keyword, file.name):
            get_file(file.name, folder)

if __name__ == '__main__':
    if FILE_NAME != 'None':
        get_file(FILE_NAME, FOLDER_NAME)
    elif FILE_NAME_PATTERN != 'None':
        get_files_by_pattern(FILE_NAME_PATTERN, FOLDER_NAME)
    else:
        get_files(FOLDER_NAME)