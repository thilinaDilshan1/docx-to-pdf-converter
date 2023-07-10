import os
from docx import Document
import win32com.client as win32

# Define the paths for your Word document and the desired PDF output
word_file = "D:/MyDoc/Cybersecurity.docx"
pdf_file = "D:/MyDoc/Cybersecurity.pdf"

# Open the Word document using pywin32
word = win32.gencache.EnsureDispatch('Word.Application')
doc = word.Documents.Open(word_file)

# Save the Word document as a PDF
doc.SaveAs(pdf_file, FileFormat=17)  # FileFormat=17 is for PDF
doc.Close()
word.Quit()
