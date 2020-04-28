import cgi
import webbrowser
import glob
import win32com.client
import os
import datetime
import sys
print("Python version")
print(sys.version)
print("Version info.")
print(sys.version_info)

word = win32com.client.Dispatch("Word.Application")
word.visible = 0
txtDir = "C:/Users/" # dir to save archive(exe: "C:/Users/")
url ='https://archive.pdf' #url get archive
pdfs_path = "" # folder where the .pdf files are stored
#for i, doc in enumerate(glob.iglob(pdfs_path+"*.pdf")): if more archive
doc =url
print('acess o archive')
print(doc)
filename = doc.split('/')[-1]
in_file = os.path.abspath(doc)
print('in file:'+in_file)
wb = word.Documents.Open(url)
print(wb)
out_file = os.path.abspath(txtDir +filename[0:-4]+ ".docx".format()) #extension archive

print("outfile\n",out_file)

wb.SaveAs2(out_file, FileFormat=16) # cod format archive

print("success...")
wb.Close()

word.Quit()


