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

    try:
        path = f"{xlLoc}\\{fileName}"
        print(path)
        assert os.path.isfile(path)

        # Load the workbook
        currentWB = openpyxl.load_workbook(f"{xlLoc}\\{fileName}") 

    # Error checking for file path exceptions
    except AssertionError as a:
        print(a)
    except InvalidFileException as e:
        print(e)

    # Load the invoices worksheet
    currentWS = currentWB['Invoices']

    # Load all headings into an array for ease of use
    wsHeadings = []
    print('Total number of rows: '+str(currentWS.max_row)+'. And total number of columns: '+str(currentWS.max_column))
    for i in range(1, currentWS.max_column+1):
        wsHeadings.append(currentWS.cell(row = 1, column=i).value)
    
    # Locate SD Document row -> Shouldnt change name as SAP is consistent across all fields
    
main()
