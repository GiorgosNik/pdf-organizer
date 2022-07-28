import os
import PyPDF2


def main():
    root = input('Please enter the root directory: ')
    # root = "E:\Documents\GitHub\pdf-organizer\TEST"

    while not os.path.isdir(root):
        root = input('The directory does not exist. Please try again: ')

    results = []

    # Get all files in directory
    fileList = []
    for root, dirs, files in os.walk(root):
        for file in files:
            # append the file name to the list
            fileList.append(os.path.join(root, file))

    terms = input('Please enter the search terms, separated by comma: ')
    terms = terms.lower()
    terms = terms.split(",")
    for i in range(len(terms)):
        terms[i] = terms[i].strip()

    #  Check each file
    for file in fileList:
        if file.endswith(".pdf"):
            # Open file
            pdfFileObj = open(file, 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj,  strict=False)

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
            for term in terms:
                if term not in accumulation:
                    notFound = True
                    break
            if not notFound:
                results.append(file)

    for i in range(len(results)):
        results[i] = results[i].replace("\\\\", "\\")

    if len(results) == 0:
        print("No files match this criteria")
    else:
        print("Results: ")
        for result in results:
            print(result)

if __name__ == '__main__':
    main()
