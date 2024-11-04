from pdf2docx import parse
import os

no_of_files = int(input("enter the number of files? "))

for count in range(no_of_files):
    file = input("enter your file path") # when u enter the path of the file remove the qoutation marks
    word_doc = f"sample{count}.docx"

    parse(file,word_doc)

    if os.path.exists(f"sample{count}.docx"):
        print("file created sucessfully")
    else:
        print("file not created")

    count += 1