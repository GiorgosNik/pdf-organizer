import PyPDF2
import os
import tkinter.filedialog
import threading
from tkinter import ttk
from tkinter import *
from ctypes import windll
import tkinter.scrolledtext as st
import multiprocessing
import numpy
import xlsxwriter
import pandas as pd
import re
import ctypes.wintypes

# General Use Globals
directory = ""
cancelFlag = False
customUI = None
results = []
search_type = ""

# Set the progress to be used for the progressbar
progress = 0

# Used for the custom window
lastClickX = 0
lastClickY = 0

# Threading Stuff
processNum = 8
completedDocuments = 0
numberOfDocuments = 0


def searchMain(givenTerms):
    global processNum
    global directory
    global customUI
    global cancelFlag
    global results
    global progress
    global completedDocuments
    global numberOfDocuments
    global search_type

    queue = multiprocessing.Queue()

    resultCounter = 0
    completedDocuments = 0
    results = []
    cancelFlag = False
    search_type = "Body"

    # Print errors
    if givenTerms == "":
        customUI.set_results("Please give at least one term")
        return

    if directory == "":
        customUI.set_results("Please select a folder")
        return

    while not os.path.isdir(directory):
        customUI.set_results("The directory does not exist. Please try again")
        return

    # Get all files in directory
    fileList = []
    temp_directory = directory
    for temp_directory, dirs, givenFiles in os.walk(temp_directory):
        for file in givenFiles:
            # Append the file name to the list
            if file.endswith(".pdf"):
                fileList.append(os.path.join(temp_directory, file))

    # Perform formatting on the search terms
    givenTerms = givenTerms.lower()
    termsToPrint = givenTerms
    givenTerms = givenTerms.split(",")
    for i in range(len(givenTerms)):
        givenTerms[i] = givenTerms[i].strip()

    # Set the number of documents globally to track progress
    numberOfDocuments = len(fileList)

    # Split the fileList
    fileList = numpy.array(fileList)
    fileList = numpy.array_split(fileList, processNum)
    for i in range(len(fileList)):
        fileList[i] = list(fileList[i])

    # Create and join processes
    process = [None] * processNum
    for i in range(processNum):
        process[i] = multiprocessing.Process(target=searchThread, args=(givenTerms, fileList[i], queue))
        process[i].start()

    # Process monitoring loop
    while 1 == 1:

        # Check for cancellation command
        if cancelFlag:
            for i in range(processNum):
                process[i].kill()
                process[i].join()
            cancelFlag = False
            progress = 0
            customUI.set_results("Search Canceled")
            return

        # Check the content of the queue
        ret = queue.get()

        if type(ret) == str:
            # If we got a string, count as progress
            completedDocuments += 1
            progress = (completedDocuments / numberOfDocuments) * 100
        elif isinstance(ret, list):
            # If we got a list, count a process as finished
            results.extend(ret)
            resultCounter += 1
            if resultCounter == processNum:
                break

    # Join all processes
    for i in range(processNum):
        process[i].join()

    # Perform formatting on the results
    for i in range(len(results)):
        results[i] = results[i].replace("\\\\", "\\")

    # Print Results
    if len(results) == 0:
        customUI.set_results("")
        customUI.set_results("No files match this criteria")
    else:
        progress = 100
        formattedResults = "Results for: " + termsToPrint + "\n"
        for result in results:
            formattedResults += result + "\n"
        customUI.set_results(formattedResults)


def searchThread(givenTerms, givenFiles, queue):
    localResults = []

    #  Check each file
    for file in givenFiles:
        try:

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

            # Convert the text to lowercase
            accumulation = accumulation.lower()

            # Remove Punctuation
            accumulation = re.sub(r'[^\w\s]', '', accumulation)

            # Check contents
            notFound = False
            for term in givenTerms:
                term = " " + term
                if term not in accumulation:
                    notFound = True
                    break
            if not notFound:
                localResults.append(file)
            queue.put("Completed One")
        except:
            # If we cant open the file, move on
            print("Could not open file:" + file)
            queue.put("Completed One")
            continue
    queue.put(localResults)


def search_title(givenTerms):
    global directory
    global customUI
    global progress
    global cancelFlag
    global results
    global search_type

    search_type = "Title"
    results = []
    cancelFlag = False

    # Display error messages
    if givenTerms == "":
        customUI.set_results("Please give at least one term")
        return

    if directory == "":
        customUI.set_results("Please select a folder")
        return

    while not os.path.isdir(directory):
        customUI.set_results("The directory does not exist. Please try again")
        return

    # Get all files in directory
    fileList = []
    temp_directory = directory
    for temp_directory, dirs, files in os.walk(temp_directory):
        for file in files:
            # Append the file name to the list
            fileList.append(os.path.join(temp_directory, file))

    # Format the search terms
    givenTerms = givenTerms.lower()
    termsToPrint = givenTerms
    givenTerms = givenTerms.split(",")
    for i in range(len(givenTerms)):
        givenTerms[i] = givenTerms[i].strip()

    #  Check each file
    index = 0
    for file in fileList:
        if cancelFlag:
            cancelFlag = False
            customUI.set_results("Search Canceled")
            progress = 0
            return

        # Check if the file is a pdf
        if file.endswith(".pdf"):
            accumulation = file.split("\\")[-1]
            accumulation = accumulation.lower()

            # Check if the title matches the search terms
            notFound = False
            for term in givenTerms:
                if term not in accumulation:
                    notFound = True
                    break
            if not notFound:
                results.append(file)

        # Count progress
        index += 1
        progress = (index / len(fileList)) * 100

    # Format the results
    for i in range(len(results)):
        results[i] = results[i].replace("\\\\", "\\")

    # Print the results
    if len(results) == 0:
        customUI.set_results("")
        customUI.set_results("No files match this criteria")
    else:
        progress = 100
        formattedResults = "Results for: " + termsToPrint + "\n"
        for result in results:
            formattedResults += result + "\n"
            customUI.set_results(formattedResults)

def getDocumentsDirectory():
    CSIDL_PERSONAL = 5  # My Documents
    SHGFP_TYPE_CURRENT = 0  # Get current, not default value

    buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
    ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)

    return buf.value

def browse_button():
    # Handler for the directory selection button
    global directory
    filename = tkinter.filedialog.askdirectory()
    directory = filename


def save_to_sheet():
    # Function to save the results in Excel-compatible format

    global search_type
    global directory
    global customUI

    # Retrieve the results
    toSave = customUI.retrieve_results()

    if len(toSave) == 0 or toSave == "Please select a folder" or toSave == "No files match this criteria" or toSave == "Search Canceled":
        customUI.set_results("Cant save empty results")
        return

    # Split the results
    toSave = toSave.splitlines()

    for i in range(1, len(toSave)):
        toSave[i] = os.path.abspath(toSave[i])

    previous = []
    if os.path.exists("organizer_results.xlsx"):
        df = pd.read_excel(r'organizer_results.xlsx')
        previous = df.values.tolist()

    documents = getDocumentsDirectory()
    if not os.path.exists(documents+"/Organizer"):
        os.makedirs(documents+"/Organizer")

    # Create and format the workbook
    workbook = xlsxwriter.Workbook(documents+"/Organizer/organizer_results.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 50)
    worksheet.set_column('D:D', 70)
    worksheet.set_column('E:E', 150)

    #  Get the search term
    searchTerms = toSave[0].split(":")[1].strip()

    toSave.pop(0)
    worksheet.write(0, 0, "TERM")
    worksheet.write(0, 1, "TYPE")
    worksheet.write(0, 2, "ROOT DIRECTORY")
    worksheet.write(0, 3, "LOCATION")
    worksheet.write(0, 4, "RESULT")

    for result in toSave:
        previous.append([searchTerms, search_type, directory, result.rsplit('\\', 1)[0], result.rsplit('\\', 1)[1]])

    toWrite = []
    for result in previous:
        flag = True
        for deduplicatedResult in toWrite:
            if len(set(result) & set(deduplicatedResult)) == 5:
                flag = False
                break
        if flag:
            toWrite.append(result)

    for result in toWrite:
        print(result)

    for i in range(len(toWrite)):
        worksheet.write(i + 1, 0, toWrite[i][0])
        worksheet.write(i + 1, 1, toWrite[i][1])
        worksheet.write(i + 1, 2, toWrite[i][2])
        worksheet.write(i + 1, 3, toWrite[i][3])
        worksheet.write(i + 1, 4, toWrite[i][4])

    workbook.close()

    customUI.set_results("Results Saved at: "+ documents+"/Organizer/organizer_results.xlsx")


def search_body_button_click(given_terms):
    # On click function to start and pass argument to the search thread
    given_terms = given_terms.get()
    process_thread = threading.Thread(target=searchMain, args=(given_terms,))
    process_thread.start()


def search_title_button_click(given_terms):
    # On click function to start and pass argument to the search thread
    given_terms = given_terms.get()
    process_thread = threading.Thread(target=search_title, args=(given_terms,))
    process_thread.start()


def cancel_search_button_click():
    # Set the global cancellation flag
    global cancelFlag
    cancelFlag = True


def set_app_window(root):
    # Needed for custom UI
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
    # Needed for custom UI
    global lastClickX, lastClickY
    lastClickX = event.x
    lastClickY = event.y


class UI:
    def __init__(self, master):
        self.master = master
        master.overrideredirect(True)  # turns off title bar, geometry
        master.geometry('950x345+700+400')  # set new geometry
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
        self.E1 = ttk.Entry(self.master, width=35)
        self.window.create_window(120, 15, anchor='nw', window=self.E1)
        self.search_terms = tkinter.Label(self.master, text="Search Terms")
        self.window.create_window(10, 20, anchor='nw', window=self.search_terms)

        self.entryString = self.E1.get()

        self.v = tkinter.StringVar()

        self.save_results = ttk.Button(self.master, text="Save Results", style='Accent.TButton',
                                       command=save_to_sheet)
        self.window.create_window(400, 15, anchor='nw', window=self.save_results)

        self.select_folder = ttk.Button(self.master, text="Select folder", style='Accent.TButton',
                                        command=browse_button)
        self.window.create_window(510, 15, anchor='nw', window=self.select_folder)

        self.search_button = ttk.Button(self.master, text="Search Body", style='Accent.TButton',
                                        command=lambda: search_body_button_click(self.E1))
        self.window.create_window(840, 15, anchor='nw', window=self.search_button)

        self.cancel_button = ttk.Button(self.master, text="Stop Search", style='Accent.TButton',
                                        command=lambda: cancel_search_button_click())
        self.window.create_window(620, 15, anchor='nw', window=self.cancel_button)

        self.search_button = ttk.Button(self.master, text="Search Title", style='Accent.TButton',
                                        command=lambda: search_title_button_click(self.E1))
        self.window.create_window(730, 15, anchor='nw', window=self.search_button)

        self.progress_label = ttk.Label(self.master, text="Progress:")
        self.window.create_window(10, 60, anchor='nw', window=self.progress_label)

        self.progress_bar = ttk.Progressbar(self.master, orient='horizontal', length=820, mode='determinate')
        self.window.create_window(120, 65, anchor='nw', window=self.progress_bar)

        self.text_area = st.ScrolledText(self.master, width=153, height=15, font=("Times New Roman", 10))
        self.window.create_window(10, 85, anchor='nw', window=self.text_area)

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

    def retrieve_results(self):
        return self.text_area.get("1.0", 'end-1c')


def main():
    global customUI

    # Set the UI as a global, start mainloop
    root = Tk()
    customUI = UI(root)

    customUI.set_progress()
    root.after(30, set_app_window, root)
    root.mainloop()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()
