import pytesseract
from PIL import Image
from pdf2image import convert_from_path
from docx import Document
import os
from tkinter import *
from tkinter import ttk


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
poppler_path = r"C:\Users\ACER\AppData\Local\Programs\VS Code\vs code- python\pdf converters\pdf_image to word\poppler-24.08.0\Library\bin"

no_of_files = int(input("enter how many pdf do u have ?"))
count = 0
for count in range(no_of_files):
    file = (input("enter file path !!")) # when u enter the path of the file remove the qoutation marks
    content = rf"{file}"
    
    images = convert_from_path(
        pdf_path = file  ,
        poppler_path=poppler_path,
    )

    text = " "
    for image in images:
        text += pytesseract.image_to_string(image)

    document = Document()
    document.add_paragraph(text)
    document.save(f"sample{count}.docx")

    if os.path.exists(f"sample{count}.docx"):
        print("Document created successfully.")
    else:
        print("Document was not created.")

    count += 1

window = Tk()
window.title("PDF to Word Converter")

window.columnconfigure(0 , weight=1)
window.columnconfigure(1, weight=1)

path_label = Label(window, text= "no of files")
path_label.grid(column=0, row=1)
path_label = Entry()



