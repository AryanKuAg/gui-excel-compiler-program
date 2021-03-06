import webbrowser
from tkinter import *
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import xlrd
from openpyxl import Workbook

from allFilesJointer import allFilesJointer

no_of_files_selected = 0

#################################

#################################
excelFileNames = ''


def select_files():
    filetypes = (
        ('Excel Files', '*.xlsx'),
        ('All files', '*.*')
    )

    filenames = fd.askopenfilenames(
        title='Open files',
        initialdir='/',
        filetypes=filetypes)

    no_of_files_selected = len(filenames)
    if no_of_files_selected > 0:
        extraction_button_updater()
        update_selected_file_status(no_of_files_selected)
        global excelFileNames
        excelFileNames = filenames


######################################


root = Tk()
root.title('Multiple Excel Files Joiner by Alemantrix Aryan Agrawal')
# it contain the app icon
root.iconbitmap('.\icon.ico')
root.resizable(False, False)
root.geometry('300x300')

space1 = Label(root, text=''' 
''').pack(
)  # This is for the space between two buttons

select_file_button = Button(
    root, text="Select Files", padx=60, pady=20, command=select_files).pack()
selected_files_status = Label(root, text="No Files Selected")
selected_files_status.pack()


def update_selected_file_status(no):
    global selected_files_status
    selected_files_status.config(text=str(no) + ' files are selected')


##########


def ask_the_directory():
    print('ask teh directory')
    # directory = fd.SaveFileDialog(master='/')
    savefilename = fd.asksaveasfilename(
        initialdir='/', title='Save File', filetypes=(('Excel File', 'Compiled.xlsx'), ('All Files', '*.*')))

    if savefilename and excelFileNames != '':
        allFilesJointer(excelFileNames, savefilename)
        extract_button.config(state=DISABLED)


# extraction_location_button = Button(
#     root, text="Select Location", padx=50, pady=20, command=ask_the_directory).pack()
# extraction_location_status = Label(root, text="No Location Selected")
# extraction_location_status.pack()
# ##########


####################
space2 = Label(root, text='''
MERA PAISA BAAKI HAI

''').pack()

##########
# global extract_button
extract_button = Button(root, text="Extract", padx=30,
                        pady=10, command=ask_the_directory)

if no_of_files_selected == 0:
    # extract_button = Button(root, text="Extract", padx=30,
    #                         pady=10, state=DISABLED)
    extract_button.config(state=DISABLED)


extract_button.pack()
extraction_status = Label(root, text="Nothing Happening")

##########


def extraction_button_updater():
    global extract_button
    extract_button.config(state=ACTIVE)


space3 = Label(root, text='''''').pack()


def follow_me():
    url = 'https://www.instagram.com/alemantrixaryanagrawal/'
    chrome_path = 'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s'
    webbrowser.get(chrome_path).open(url)


Button(root, text="Follow Me!!!", command=lambda: follow_me()).pack()
Label(text="Software By ?? 2022 Alemantrix").pack()

root.mainloop()
