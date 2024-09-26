from plistlib import InvalidFileException
from pdfminer.high_level import extract_pages, extract_text
import pypdf, openpyxl, os, re
from helpers import *


def main():
    # User provides the location of the folders where invoices are
    folLoc = input("Provide location of Folders: ")
    folLoc = str(checkIfBlank(folLoc,"location"))
    
    # Continue through all pdf files in folder to output the sum of each type of fee
    for file in os.listdir(folLoc):
        filename = os.fsdecode(file)
        
        # Check for pdf
        if filename.endswith(".pdf"):
            print(file)
            pdfText = extract_text(f'{folLoc}\\{filename}')

            # series of lists for the position and sum of each type of fee per invoice
            esFee = sum([float(pdfText[x.end():x.end()+4]) for x in re.finditer("Establishment Fee", pdfText)])
            print(esFee)
            '''
            for page in extract_pages(f'{folLoc}\\{filename}'):
                for element in page:
                    print(element)
            '''

        else:
            print(f'{filename} is not a pdf')


main()