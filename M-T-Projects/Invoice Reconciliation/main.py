from plistlib import InvalidFileException
from pdfminer.high_level import extract_pages, extract_text
import pypdf, openpyxl, os, re
from helpers import *


def main():
    # User provides the location of the folders where invoices are
    folLoc = input("Provide location of Folders: ")
    folLoc = str(checkIfBlank(folLoc,"location"))
    
    
    for file in os.listdir(folLoc):
        filename = os.fsdecode(file)
        if filename.endswith(".pdf"):
            print(file)
            for page in extract_pages(f'{folLoc}\\{filename}'):
                for element in page:
                    print(element)


        else:
            print(f'{filename} is not a pdf')


main()