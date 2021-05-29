import os
import glob
from PyPDF2 import PdfFileWriter, PdfFileMerger, PdfFileReader
from PIL import Image
import pytesseract
import fitz
from shutil import copyfile

#empty file and initial environment
files = glob.glob(r'C:\Users\pisce\Downloads\BatchLetter\Split\*')
for f in files:
    os.remove(f)
files = glob.glob(r'C:\Users\pisce\Downloads\BatchLetter\Output\*')
for f in files:
    os.remove(f)

#pdf split
pdfs = os.listdir(r'C:\Users\pisce\Downloads\BatchLetter\PDF')
for p in pdfs:
    pdf = "C:\\Users\\pisce\\Downloads\\BatchLetter\\PDF\\"+str(p)
    pdf_file = PdfFileReader(open(pdf,"rb"),strict = False)
    pdf_pages_len = pdf_file.getNumPages() // 2
    # print(p[6:8]+"|"+str(pdf_pages_len))
    for n in range(pdf_pages_len):
        output = PdfFileWriter()
        start_page = n*2
        end_page = (n+1)*2
        for i in range(start_page, end_page):
            output.addPage(pdf_file.getPage(i))
            newname = 'C:\\Users\\pisce\\Downloads\\BatchLetter\\Split\\'+p[6:8]+"_"+str(n+1)+".pdf"
        outputStream = open(newname, "wb")
        output.write(outputStream)
        outputStream.close()
        # print('File:'+p[6:8]+' |Start Page:'+str(start_page)+' |End Page:'+str(end_page))

#tesseract cmd initial
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

#ocr and file rename
fdir = os.listdir('C:\\Users\\pisce\\Downloads\\BatchLetter\\Split\\')
for i in range(len(fdir)):    
    fname="C:\\Users\\pisce\\Downloads\\BatchLetter\\Split\\"+str(fdir[i])
    pdffile = fname
    doc = fitz.open(pdffile)
    page = doc.loadPage(0)
    zoom = 2    # zoom factor
    mat = fitz.Matrix(zoom, zoom)
    pix = page.getPixmap(matrix=fitz.Matrix(150/72,150/72))
    output = "C:\\Users\\pisce\\Downloads\\BatchLetter\\test.png"
    pix.writePNG(output)
    im = Image.open(output) 
    left = 1000
    right = 1115
    top = 1450
    bottom = 1510
    im1 = im.crop((left, top, right, bottom)) 
    words = pytesseract.image_to_string(im1)
    words2 = words.split("\n") 
    fname2 = fname[-9:-7]+'_'+words2[0][3:8]
    src=fname
    dst="C:\\Users\\pisce\\Downloads\\BatchLetter\\Output\\"+fname2+".pdf"
    copyfile(src, dst)
    print("File No."+str(i+1)+"|File Name: "+fname2+".pdf")
    i+1