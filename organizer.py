import os
import PyPDF2
import os
import glob
import requests
import shutil
import random
import tkinter.filedialog
import threading
from tkinter import ttk
from zipfile import ZipFile
from tkinter import *
from ctypes import windll
import time
import tkinter.scrolledtext as st

directory = ""
terms = ""
customUI = None
# Set the progress to be used for the progressbar
progress = 0

# Used for the custom window
lastClickX = 0
lastClickY = 0

# Used by tkinter

def search_body(givenTerms):
    global directory
    global customUI
    global progress
    results = []

    if givenTerms == "":
        customUI.spawn_message("Please give at least one term")
        return

    if directory == "":
        customUI.spawn_message("Please select a folder")
        return

    while not os.path.isdir(directory):
        customUI.spawn_message("The directory does not exist. Please try again")
        return

    # Get all files in directory
    fileList = []
    temp_directory = directory
    for temp_directory, dirs, files in os.walk(temp_directory):
        for file in files:
            # append the file name to the list
            fileList.append(os.path.join(temp_directory, file))

    givenTerms = givenTerms.lower()
    givenTerms = givenTerms.split(",")
    for i in range(len(givenTerms)):
        givenTerms[i] = givenTerms[i].strip()

    #  Check each file
    index = 0
    for file in fileList:
        if file.endswith(".pdf"):
            # Open file
            pdfFileObj = open(file, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

            # Accumulate text from each page
            pageNum = pdfReader.numPages
            accumulation = ""
            for page in range(0, pageNum):
                pageObj = pdfReader.getPage(page)
                accumulation += pageObj.extractText()

            # Close the file
            pdfFileObj.close()

            accumulation = accumulation.lower()

            # Check contents
            notFound = False
            for term in givenTerms:
                if term not in accumulation:
                    notFound = True
                    break
            if not notFound:
                results.append(file)
        index+=1
        progress=(index/len(fileList))*100
        customUI.set_progress()
    for i in range(len(results)):
        results[i] = results[i].replace("\\\\", "\\")

    if len(results) == 0:
        customUI.set_results("")
        customUI.spawn_message("No files match this criteria")
    else:
        formatedResults=""
        for result in results:
            formatedResults+=result+"\n"
        customUI.set_results(formatedResults)

def search_title(givenTerms):
    global directory
    global customUI
    global progress
    results = []

    if givenTerms == "":
        customUI.spawn_message("Please give at least one term")
        return

    if directory == "":
        customUI.spawn_message("Please select a folder")
        return

    while not os.path.isdir(directory):
        customUI.spawn_message("The directory does not exist. Please try again")
        return

    # Get all files in directory
    fileList = []
    temp_directory = directory
    for temp_directory, dirs, files in os.walk(temp_directory):
        for file in files:
            # append the file name to the list
            fileList.append(os.path.join(temp_directory, file))

    givenTerms = givenTerms.lower()
    givenTerms = givenTerms.split(",")
    for i in range(len(givenTerms)):
        givenTerms[i] = givenTerms[i].strip()

    #  Check each file
    index = 0
    for file in fileList:
        if file.endswith(".pdf"):

            accumulation = file.split("\\")[-1]
            accumulation = accumulation.lower()

            # Check contents
            notFound = False
            for term in givenTerms:
                if term not in accumulation:
                    notFound = True
                    break
            if not notFound:
                results.append(file)
        index+=1
        progress=(index/len(fileList))*100
        customUI.set_progress()
    for i in range(len(results)):
        results[i] = results[i].replace("\\\\", "\\")

    if len(results) == 0:
        customUI.set_results("")
        customUI.spawn_message("No files match this criteria")
    else:
        formattedResults=""
        for result in results:
            formattedResults+=result+"\n"
        customUI.set_results(formattedResults)

def browse_button():
    global directory
    filename = tkinter.filedialog.askdirectory()
    directory = filename


def search_body_button_click(given_terms):
    given_terms = given_terms.get()
    process_thread = threading.Thread(target=search_body, args=(given_terms,))
    process_thread.start()

def search_title_button_click(given_terms):
    given_terms = given_terms.get()
    process_thread = threading.Thread(target=search_title, args=(given_terms,))
    process_thread.start()

def set_app_window(root):
    GWL_EXSTYLE = -20
    WS_EX_APPWINDOW = 0x00040000
    WS_EX_TOOLWINDOW = 0x00000080
    hwnd = windll.user32.GetParent(root.winfo_id())
    style = windll.user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
    style = style & ~WS_EX_TOOLWINDOW
    style = style | WS_EX_APPWINDOW
    res = windll.user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, style)
    # re-assert the new window style
    root.withdraw()
    root.after(10, root.deiconify)

def save_last_click_pos(event):
    global lastClickX, lastClickY
    lastClickX = event.x
    lastClickY = event.y

class UI:
    def __init__(self, master):
        self.master = master
        master.overrideredirect(True)  # turns off title bar, geometry
        master.geometry('500x390+700+400')  # set new geometry
        master.attributes('-topmost', False)
        master.call("source", "forest-dark.tcl")
        ttk.Style().theme_use('forest-dark')
        self.title_bar = Frame(self.master, bg='#242424', relief='groove', bd=2, highlightthickness=0)
        self.title_bar.config(borderwidth=0, highlightthickness=0)
        self.close_button = Button(self.title_bar, text='X', command=self.master.destroy, bg="#242424", padx=2, pady=2,
                                   activebackground='red',
                                   bd=0, font="bold", fg='white', highlightthickness=0)
        # Create a window
        self.window = Canvas(self.master, bg='#2e2e2e', highlightthickness=0)

        # Pack the above widgets
        self.title_bar.pack(expand=0, fill=BOTH)
        self.close_button.pack(side=RIGHT)
        self.window.pack(expand=1, fill=BOTH)
        self.title_bar.bind('<Button-1>', save_last_click_pos)
        self.title_bar.bind('<B1-Motion>', self.dragging)

        # Create the widgets
        self.master.after(10, set_app_window, self.master)
        self.E1 = ttk.Entry(self.master, width=50)
        self.window.create_window(120, 15, anchor='nw', window=self.E1)
        self.playlist_label = tkinter.Label(self.master, text="Search Terms")
        self.window.create_window(10, 20, anchor='nw', window=self.playlist_label)

        self.entryString = self.E1.get()

        self.v = tkinter.StringVar()

        self.select_folder = ttk.Button(self.master, text="Select folder", style='Accent.TButton',
                                        command=browse_button)
        self.window.create_window(280, 60, anchor='nw', window=self.select_folder)

        self.search_button = ttk.Button(self.master, text="Search Body", style='Accent.TButton',
                                          command=lambda: search_body_button_click(self.E1))
        self.window.create_window(390, 60, anchor='nw', window=self.search_button)

        self.search_button = ttk.Button(self.master, text="Search Title", style='Accent.TButton',
                                        command=lambda: search_title_button_click(self.E1))
        self.window.create_window(170, 60, anchor='nw', window=self.search_button)

        self.progress_label = ttk.Label(self.master, text="Progress:")
        self.window.create_window(10, 105, anchor='nw', window=self.progress_label)

        self.progress_bar = ttk.Progressbar(self.master, orient='horizontal', length=365, mode='determinate')
        self.window.create_window(120, 110, anchor='nw', window=self.progress_bar)

        self.text_area = st.ScrolledText(self.master,
                                    width=76,
                                    height=15,
                                    font=("Times New Roman",
                                          10))

        self.window.create_window(10, 130, anchor='nw', window=self.text_area)

        # Making the text read only
        self.text_area.configure(state='disabled')

        self.messageBox = None

    def set_results(self, results):
        self.text_area.configure(state='normal')
        self.text_area.delete(1.0, tkinter.END)
        self.text_area.insert(tkinter.INSERT, results)
        self.text_area.configure(state='disabled')

    def set_progress(self):
        global progress
        self.progress_bar['value'] = progress
        self.master.after(60, self.set_progress)

    def dragging(self, event):
        x, y = event.x - lastClickX + self.master.winfo_x(), event.y - lastClickY + self.master.winfo_y()
        self.master.geometry("+%s+%s" % (x, y))

    def spawn_message(self, given_message, given_title=None):
        self.messageBox = tkinter.messagebox.showinfo(title=given_title, message=given_message)
    def spawn_results(self, given_message, given_title=None):
        self.messageBox = tkinter.messagebox.showinfo(title=given_title, message=given_message)

def main():
    global directory
    global terms
    global customUI
    # # root = "E:\Downloads\Organic"

    root = Tk()
    customUI = UI(root)

    terms = customUI.E1.get()
    customUI.set_progress()
    root.after(1, set_app_window, root)
    root.mainloop()


if __name__ == '__main__':
    main()
