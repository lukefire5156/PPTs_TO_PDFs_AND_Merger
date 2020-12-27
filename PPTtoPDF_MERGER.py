#lukefire
import sys
import os
import comtypes.client
import time
from PyPDF2 import PdfFileMerger
import re

inputFolder = sys.argv[1]  #input folder path that contains ppts to convert
outputFolder = inputFolder

inputFolder = os.path.abspath(inputFolder)
outputFolder = os.path.abspath(outputFolder)

AllPPTFileNames = os.listdir(inputFolder)   #get path of all the ppt files in the folder

print("Converting PPTs to PDFs:")

animation = ["[■□□□□□□□□□]","[■■□□□□□□□□]", "[■■■□□□□□□□]", "[■■■■□□□□□□]", "[■■■■■□□□□□]", "[■■■■■■□□□□]", "[■■■■■■■□□□]", "[■■■■■■■■□□]", "[■■■■■■■■■□]", "[■■■■■■■■■■]"]

for i in range(len(AllPPTFileNames)):
    PPTName=AllPPTFileNames[i]
    if not PPTName.lower().endswith((".ppt", ".pptx")):
        continue
    
    PPTPath = os.path.join(inputFolder, PPTName)
        
    PPT = comtypes.client.CreateObject("Powerpoint.Application")
    
    PPT.Visible = 1
    
    slides = PPT.Presentations.Open(PPTPath)
    
    file_name = os.path.splitext(PPTName)[0]
    
    output_file_path = os.path.join(outputFolder, file_name + ".pdf")
    
    slides.SaveAs(output_file_path, 32)

    slides.Close()
    time.sleep(0.2)
    x=int(((i+1)/len(AllPPTFileNames))*10)
    if x>0:
        x-=1
    sys.stdout.write("\r" + animation[x])
    sys.stdout.flush()

print("\nALL PPTs Converted to PDFs Successfully!")
def alphaNumOrder(string):
    return ''.join([format(int(x), '05d') if x.isdigit() else x for x in re.split(r'(\d+)', string)])

print("Merging PDFs:")
    
merger = PdfFileMerger()
AllFiles = os.listdir(inputFolder)
AllFiles.sort(key=alphaNumOrder) 
#print(AllFiles)
for item in range(len(AllFiles)):
    items=AllFiles[item]
    if items.endswith('.pdf'):
        merger.append(os.path.join(inputFolder, items))
    x=int(((item+1)/len(AllFiles))*10)
    if x>0:
        x-=1
    sys.stdout.write("\r" + animation[x])
    sys.stdout.flush()		


merger.write(os.path.join(outputFolder,"MergedPDF.pdf"))
merger.close()
print("\nPDFs Merged Successfully!")
