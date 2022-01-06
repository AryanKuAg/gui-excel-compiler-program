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
        allFilesJointer(filenames)

######################################


root = Tk()
root.title('Multiple Excel Files Joiner by Alemantrix Aryan Agrawal')
root.iconbitmap()  # it contain the app icon
root.resizable(False, False)
root.geometry('300x350')

space1 = Label(root, text=''' ''').pack(
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
    directory = fd.askdirectory()
    extraction_location_status.config(text=directory)
    print(directory)


extraction_location_button = Button(
    root, text="Select Location", padx=50, pady=20, command=ask_the_directory).pack()
extraction_location_status = Label(root, text="No Location Selected")
extraction_location_status.pack()
##########

space2 = Label(root, text='''       


''').pack()

##########
# global extract_button
extract_button = Button(root, text="Extract", padx=30,
                        pady=10, command=ask_the_directory)

if no_of_files_selected == 0:
    extract_button = Button(root, text="Extract", padx=30,
                            pady=10, state=DISABLED)


extract_button.pack()
extraction_status = Label(root, text="Nothing Happening")

##########


def extraction_button_updater():
    global extract_button
    extract_button.config(state=ACTIVE)
    # extract_button.pack_forget()
    # extract_button = Button(root, text="Extract", padx=30,
    #                         pady=10).pack()


space3 = Label(root, text='''       


''').pack()

root.mainloop()
