# Python Business Process Automation
Python Business Automation is mainly focusing on using business process automation via python. By automating business process, reducing manual processes and increasing efficiency and accuracy. 

---
## 1.PDF-Split-and-Rename

#### The PDF information page extract is using the Tesseract. 

This Python Notebook is for PDF invoice processing, which includes Split the PDF into invididual file and Rename the PDF file using with information on first page


----------------------------------------------------

#### The initialization guide lists below:
1. Install tesseract using windows installer available at: https://github.com/UB-Mannheim/tesseract/wiki

2. Note the tesseract path from the installation. Default installation path at the time of this edit was: C:\Users\USER\AppData\Local\Tesseract-OCR.
It may change so please check the installation path.

3. pip install pytesseract

4. Set the tesseract path in the script before calling image_to_string:
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\USER\AppData\Local\Tesseract-OCR\tesseract.exe'


----------------------------------------------------
#### PDF file error handle

```
PdfReadWarning: Invalid stream (index 0) within object 62 0: Stream has ended unexpectedly [pdf.py:1128]
```
When in strict mode, PyPDF2 quits when encountering this stream error and throws a PdfReadError. When strict is False, it ignores this error but instead gives a warning like you saw (then continues with rest of program as normal).

```
input = PdfFileReader([your file], strict = False)
```

----------------------------------------------------
#### Python Library Error handle (import fitz)

```
File "/home/malcolm/venvs/p3/lib/python3.8/site-packages/starlette/staticfiles.py", line 55, in __init__
    raise RuntimeError(f"Directory '{directory}' does not exist")
```
The error lines quoted from __init__.py are not contained in PyMuPDF. They demonstrate that you have installed a package named fitz in the same Python where PyMuPDF resides. This cannot coexist with PyMuPDF which has a top-level name of fitz as well. The Solution is updating the python library pymupdf.

```
pip uninstall fitz
pip install pymupdf
```


---
## 2.Python Excel Report Automation via Outlook

Nowdays, Excel still is a essetial and basic tool for many reports as business requests. This script to extract data from database, and insert into Excel (xlsx file). After that, format the data in Excel workbooks (highlight rows with selection criteria), and then send this report to users via Outlooks.

Preview:
![](https://github.com/malcolmyang25/python_business_process_automation/blob/main/python_excel_sample.png)
