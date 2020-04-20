#!/usr/bin/env python

import os
import requests
from pandas import read_excel
from tqdm import tqdm
import wget
import subprocess
import sys

def openfile(file):
    if sys.platform == 'linux2':
        subprocess.call(["xdg-open", file])
    else:
        os.startfile(file)

# insert here the folder you want the books to be downloaded:
folder = os.getcwd() + '/downloads/'

if not os.path.exists(folder):
    os.mkdir(folder)

xlsfile = "Free+English+textbooks.xlsx"

xlsname = os.path.join(os.getcwd(), xlsfile)
urldownload = 'https://resource-cms.springernature.com/springer-cms/rest/v1/content/17858272/data/v4'

if os.path.exists(xlsname):
    print(f"********* Se descargarán los libros descriptos en {xlsfile} *********")
else:
    wget.download(urldownload, xlsname)
    print("")
    print("********* Elimine las filas con los libros que no quiera, guarde y cierre la planilla*********")
    openfile(xlsname)
    #subprocess.call(f'cmd /c "start /WAIT {xlsname}')

books = read_excel(xlsname)

# save table:
books.to_excel(folder + 'table.xlsx')

# debug:
# books = books.head()

print('Download started.')

for url, title, author, pk_name in tqdm(books[['OpenURL', 'Book Title', 'Author', 'English Package Name']].values):

    new_folder = folder + pk_name + '/'

    if not os.path.exists(new_folder):
        os.mkdir(new_folder)

    r = requests.get(url) 
    new_url = r.url

    new_url = new_url.replace('/book/','/content/pdf/')

    new_url = new_url.replace('%2F','/')
    new_url = new_url + '.pdf'

    final = new_url.split('/')[-1]
    final = title.replace(',','-').replace('.','').replace('/',' ').replace(':',' ') + ' - ' + author.replace(',','-').replace('.','').replace('/',' ').replace(':',' ') + ' - ' + final
    output_file = new_folder+final
    if not os.path.exists(output_file):
        myfile = requests.get(new_url, allow_redirects=True)
        try:
            open(output_file, 'wb').write(myfile.content)
        except OSError: 
            print("Error: PDF filename is appears incorrect.")
        
        #download epub version too if exists
        new_url = r.url

        new_url = new_url.replace('/book/','/download/epub/')
        new_url = new_url.replace('%2F','/')
        new_url = new_url + '.epub'

        final = new_url.split('/')[-1]
        final = title.replace(',','-').replace('.','').replace('/',' ').replace(':',' ') + ' - ' + author.replace(',','-').replace('.','').replace('/',' ').replace(':',' ') + ' - ' + final
        output_file = new_folder+final
        
        request = requests.get(new_url)
        if request.status_code == 200:
            myfile = requests.get(new_url, allow_redirects=True)
        try:
            open(output_file, 'wb').write(myfile.content)
        except OSError: 
            print("Error: EPUB filename is appears incorrect.")
            
print('Download finished.')
openfile(folder)