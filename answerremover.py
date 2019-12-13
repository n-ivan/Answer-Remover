# ERAB
# Removes answers from exam bank questions
# Author: n-ivan
# Born: 2019-12-12

from tkinter import filedialog
from tkinter import *
from docx import Document

file_path = ""

# Selects your file
def browse_files():
    global file_path
    fileName = filedialog.askopenfilename(filetypes=(
        ("Doc Files", "*.doc"), ("Docx Files", "*.docx")))
    file_path = fileName
    disp_file['text'] = fileName.split('/')[-1]
    print(f"Selected: {file_path}")

# Resets selections
def clearEntry():
    file_path = ""
    disp_file['text'] = "Select a file..."
    ansform_entry.delete(0, END)

# Does all the actual work.
def removeAnswers():
    ansform = ansform_text.get().strip()
    if ansform == "":
        print("You haven't selected an answer format.")
    if file_path == "":
        print('You haven\'t selected a file.')
    else:
        wordDoc = Document(file_path)
        answers = []
        for para in wordDoc.paragraphs:
            try:
                if para.text.split()[0] == ansform:
                    answers.append(para.text.split()[1])
                    para.text = ''
            except:
                pass
        answerkey= open(f"{file_path.split('.')[0]}-answerkey.txt",'w')
        for i in range(len(answers)):
            answerkey.write(answers[i]+"\n")
        answerkey.close()
        wordDoc.save(f"{file_path.split('.')[0]}-no-ans.docx")
        clearEntry()

# Window Object
root = Tk()
root.title("Answer Remover")
root["bg"]="white"

# Title
title_label = Label(root, text="Exam Bank Answer Remover", font=('Avenir Next', 30))
title_label.grid(row=0,column=0)
# title_label.grid_rowconfigure(1, weight=1)
# title_label.grid_columnconfigure(1, weight=1)
title_label['bg'] = root['bg']

# Answer Format
ansform_text = StringVar()
ansform_label = Label(root, text='Answer Format (ex. Ans:)', font=('bold', 14))
ansform_label.grid(row=1, column=1)
ansform_label['bg'] = root['bg']
ansform_entry = Entry(root, textvariable=ansform_text)
ansform_entry.grid(row=1, column=2)

# File Open Button
open_file = Button(root, text="Browse Files", command=browse_files)
open_file.grid(row=2, column=0)
open_file['bg'] = root['bg']
disp_file = Label(root, text= 'Select a file...')
disp_file.grid(row=1,column=0)

# Convert Button
convert = Button(root, text="Remove Answers", command=removeAnswers)
convert.grid(row=2, column=1)
convert['bg'] = root['bg']

# Start program
root.mainloop()
