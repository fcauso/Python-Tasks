# -*- coding: utf-8 -*-
"""
@author: Fiara Causo
"""
import os
import PyPDF2
from PyPDF2 import PdfFileMerger
from PyPDF2 import  PdfFileWriter
from PyPDF2 import PdfFileReader
import os
"""
    By using os.list I obtain a list of all the files in the working directory
    the for loop will loop over the list called specifically any with an end of pdf. 
"""
#set directory of where files are
os.chdir('C:/Users/EN288JF/OneDrive - EY/Desktop/10x Innovation/ORACLE/FY21/Contract Python/TOP')
#get all the file names in the directory
all_file_names = os.listdir()
print(all_file_names)
#get specifically all pdf names
pdf_files = [f for f in os.listdir('C:/Users/EN288JF/OneDrive - EY/Desktop/10x Innovation/ORACLE/FY21/Contract Python/Top Deal') if os.path.isfile(f)]
pdf_files = filter(lambda f: f.endswith(('.pdf','.PDF')), files)
print(list(pdf_files))
"""
a lot of the files were missing the EOF Marker. 
This loop will fix the files and resave them under a new :fixed name
"""
EOF_MARKER = b'%%EOF'
file_name =  'EY RevRec Form -RR ID 90111.pdf'

with open(file_name, 'rb') as f:
    contents = f.read()

# check if EOF is somewhere else in the file
if EOF_MARKER in contents:
    # we can remove the early %%EOF and put it at the end of the file
    contents = contents.replace(EOF_MARKER, b'')
    contents = contents + EOF_MARKER
else:
    # Some files really don't have an EOF marker
    # In this case it helped to manually review the end of the file
    print(contents[-8:]) # see last characters at the end of the file
    # printed b'\n%%EO%E'
    contents = contents[:-6] + EOF_MARKER

with open(file_name.replace('.pdf', '') + '_fixed.pdf', 'wb') as f:
    f.write(contents)
"""
Ready to merge and append
"""
# Get all the PDF filenames.
pdfFiles = []
for filename in os.listdir():
    if filename.endswith('.pdf'):
        pdfFiles.append(filename)
pdfFiles.sort()
pdfWriter = PyPDF2.PdfFileWriter()
# Loop through all the PDF files.
for filename in pdfFiles:
    pdfFileObj = open(filename, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)

# Loop through all the pages and then use a slicer to only extract the last three pages and add them.
    for pageNum in range(pdfReader.numPages)[-3:]:
        pageObj = pdfReader.getPage(pageNum)
        pdfWriter.addPage(pageObj)
# Save the resulting PDF to a file.
pdfOutput = open('Q3_Oracle_TopDeals.pdf', 'wb')
pdfWriter.write(pdfOutput)
pdfOutput.close()
    