import requests
from shareplum import Office365
from config import config

#get data from configuration
username = config['sp_user']
password = config['sp_password']
site_name = config['sp_site_name']
base_path = config['sp_base_path']
doc_library = config['sp_doc_library']

file_name = s+'.pdf'