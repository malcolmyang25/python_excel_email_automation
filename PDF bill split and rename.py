import pdfplumber
import os
import fnmatch
from shutil import copyfile
from PyPDF2 import PdfFileWriter, PdfFileMerger, PdfFileReader
import glob
import sys
from PIL import Image
import pytesseract
import fitz
import re

#empty all temp files
# files = glob.glob(r'XXX\Desktop\Water Billing PDF Split\Invoice\*')
# for f in files:
#     os.remove(f)
files = glob.glob(r'XXXX\Water Billing PDF Split\SplitFiles\*')
for f in files:
    os.remove(f)

##
##set up parameter
##change the split page number
p1=8
##change the file name
pdf="XXXX\\Water Billing PDF Split\\"+str(p1)+" pages invoice.pdf"
##
pdf_file = PdfFileReader(open(pdf,"rb"),strict = False)
pdf_pages_len = pdf_file.getNumPages() // int(p1)
print(pdf_pages_len)
for n in range(pdf_pages_len):
    output = PdfFileWriter()
    start_page = n*int(p1)
    end_page = (n+1)*int(p1)
    for i in range(start_page, end_page):
        output.addPage(pdf_file.getPage(i))
        newname = r'XXXXX\Water Billing PDF Split\SplitFiles\PP'+str(p1)+" - " + str(n+1) + ".pdf"
    outputStream = open(newname, "wb")
    output.write(outputStream)
    outputStream.close()
    print('File:'+newname+' |Start Page:'+str(start_page)+' |End Page:'+str(end_page))

fname_format1="PP"+str(p1)+"*.pdf"
fdir=fnmatch.filter(os.listdir("XXXX\\Water Billing PDF Split\\SplitFiles\\"), fname_format1)
print('Total Rename File number '+str(len(fdir)))

##initial tesseract environment
pytesseract.pytesseract.tesseract_cmd = 'C:\\Users\\malcolm.yang\\AppData\\Local\\Tesseract-OCR\\tesseract.exe'

for i in range(len(fdir)):
    fname="XXXX\\Water Billing PDF Split\\SplitFiles\\"+str(fdir[i])
    pdffile = fname
    doc = fitz.open(pdffile)
    page = doc.loadPage(0)
    zoom = 2    # zoom factor
    mat = fitz.Matrix(zoom, zoom)
    pix = page.getPixmap(matrix=fitz.Matrix(150/72,150/72))
    output = "XXXX\\Desktop\\Water Billing PDF Split\\test.png"
    pix.writePNG(output)
    im = Image.open(output) 
    left = 850
    right = 1100
    top = 360
    bottom = 460
    im1 = im.crop((left, top, right, bottom)) 
    words = pytesseract.image_to_string(im1)
    words2 = words.split("\n") 
    name1_str = str(words2[0])
    name1 = int(re.search(r'\d+', name1_str)[0])
    name2_str = str(words2[2])
    name2 = int(re.search(r'\d+', name2_str)[0])
    ##quarter name needs to update
    fname2 = 'No.'+str(i+101)+'_SW_'+str(name1)+'_TC'+str(name2)+'_FY21Q2'
##
#out put directory need update
##
    if len(str(name1))<5 or len(str(name2))<5:
        src=fname
        dst="XXXX\\Water Billing PDF Split\\Discrepancy Invoice\\[Error - "+str(p1)+"p]"+fname2+".pdf"
        copyfile(src, dst)
        print("**File Error**File No."+str(i+101)+"|File Name: "+fname2+".pdf")
    else:
        src=fname
        dst="XXXX\\Water Billing PDF Split\\Invoice\\"+str(p1)+" page invoice\\"+fname2+".pdf"
        copyfile(src, dst)
        print("File No."+str(i+101)+"|File Name: "+fname2+".pdf")
    i+1