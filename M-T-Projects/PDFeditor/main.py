from plistlib import InvalidFileException
import pypdf, openpyxl, os
from helpers import *


def main():

    # User provides the location of the excel file 
    xlLoc = input("Provide location of excel file: ")
    xlLoc = str(checkIfBlank(xlLoc, "location"))

    # User provides the file to be manipulated
    fileName = input("Provide name of excel file: ")
    fileName = str(checkIfBlank(fileName, "file"))

    # Error checking for file path exceptions
    try:
        path = f"{xlLoc}\\{fileName}"
        assert os.path.isfile(path)

        # Load the workbook
        currentWB = openpyxl.load_workbook(f"{xlLoc}\\{fileName}")
    
    
    except AssertionError as a:
        print(a)
    except InvalidFileException as e:
        print(e)

    # Load the invoices worksheet
    currentWS = currentWB['Invoices']

    maxColumn = currentWS.max_column
    # Load all headings into an array for ease of use
    wsHeadings = []
    print('Total number of rows: '+str(currentWS.max_row)+'. And total number of columns: '+str(currentWS.max_column))
    for i in range(1, maxColumn+1):
        wsHeadings.append(currentWS.cell(row = 1, column=i).value)
    
    # Locate SD Document row -> Shouldnt change name as SAP is consistent across all fields
    SDColumn = wsHeadings.index("SD Document")
    invNums = []

    # The logic of the array stems from the correlation between index and row (2D)
    for i in range(1, currentWS.max_row):
        invNums.append(currentWS.cell(row=i,column=SDColumn+1).value)
    
    # Now we must find the pdf files corresponding to the invoice numbers/SD Docs
    # We will use the created array of file names (and any other fields) to conjoin them to the excel file
    pdfName = input("Please enter name of pdf to search: ")
    pdfName = str(checkIfBlank(pdfName, "name"))
    
    # Error checking for file path exceptions
    try: 
        path = f"{xlLoc}\\{fileName}"
        assert os.path.isfile(path)

        # Load the pdf
        pdfRead = pypdf.PdfReader(f"{xlLoc}\\{pdfName}")

    except AssertionError as a:
        print(a)
    except InvalidFileException as e:
        print(e)

    content = []
    # extract to textfile with the page number 
    for page in pdfRead.pages:
        pdfPageName = f"{pdfName[:-4]}_{page.page_number + 1}"

        # Maintain array if needed for more functionality
        content.append((page.extract_text(), pdfPageName))

        counter = 1
        # When the page contains the SD Document number then add that page number to the excel file 
        for i in invNums:
            if content[page.page_number][0].find(i) != -1:
                currentWS.cell(row=counter,column=maxColumn+1).value = content[page.page_number][1]
            counter += 1
        
        # Save the workbook to ensure changes
        currentWB.save(f"{xlLoc}\\{fileName}")

        # Split the PDF into multiple files
        pdfWrite = pypdf.PdfWriter()
        pdfWrite.add_page(page)

        with open(f"{xlLoc}\\{pdfPageName}.pdf", 'wb') as out:
            pdfWrite.write(out)
        print(f'Split {pdfPageName} successfully')
    
    
    
    


    
main()
