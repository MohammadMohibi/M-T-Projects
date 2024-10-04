from plistlib import InvalidFileException
from pdfminer.high_level import extract_pages, extract_text
import pypdf, openpyxl, os, re
from helpers import *

def findIteration(pdf, searchTerm):

    return sum([float(pdf[x.end():x.end()+4]) for x in re.finditer(searchTerm, pdf)])

def main():
    # file to produce the invoice results 
    writePath = r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\TL Invoice output.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # User provides the location of the folders where invoices are
    folLoc = input("Provide location of Folders: ")
    folLoc = str(checkIfBlank(folLoc,"location"))
    
    # Since invoices require two columns; 1 for volume and 2 for value we need to index +2 to match with volume only
    fileNum = 1
    # Continue through all pdf files in folder to output the sum of each type of fee
    for file in os.listdir(folLoc):
        filename = os.fsdecode(file)
        

        # Check for pdf
        if filename.endswith(".pdf"):
            print(file)
            pdfText = extract_text(f'{folLoc}\\{filename}')

            # series of lists for the position and sum of each type of fee per invoice

            # Finds the position and adds the volume of that fee in the tuple (line#, sum)
            mealFee = (58, findIteration(pdfText,r"Meal Allowance"))
            esFee = (60, findIteration(pdfText,r"Establishment Fee"))
            twoxnightFee = (12, findIteration(pdfText,r"2 x Traffic Controller Night Rate 0-8 hrs"))
            twoxnightplusFee = (24, findIteration(pdfText,r"2 x Traffic Controller Night Rate 8\+ hrs"))
            tmanightFee = (15, findIteration(pdfText,r"TMA with Operator Night Rate 0-8 hrs"))
            tmanightplusFee = (27, findIteration(pdfText,r"TMA with Operator Night Rate 8\+ hrs"))
            podFee = (52, findIteration(pdfText,r"Drop Deck / Pod Truck"))
            onexnightFee = (11, findIteration(pdfText,r"1 x Traffic Controller Night Rate 0-8 hrs"))
            onexnightplusFee = (23, findIteration(pdfText,r"1 x Traffic Controller Night Rate 8\+ hrs"))
            portFee = (55, findIteration(pdfText,r"Portaboom"))
            vmsFee = (48, findIteration(pdfText,r"VMS Ute with Operator Night Rate 0-8 hrs"))
            vmsplusFee = (49, findIteration(pdfText,r"VMS Ute with Operator Night Rate 8\+ hrs"))

            # Sort them into order of the line number so it is easier to copy
            invoiceList = (filename, sorted([mealFee,esFee,twoxnightFee, twoxnightplusFee,tmanightFee, tmanightplusFee, podFee,onexnightFee,onexnightplusFee,portFee,vmsFee,vmsplusFee]))
            print(invoiceList)

            '''
            for page in extract_pages(f'{folLoc}\\{filename}'):
                for element in page:
                    print(element)
            '''

        else:
            print(f'{filename} is not a pdf')
        
        for line in invoiceList[1]:
            sheet.cell(row=line[0],column=fileNum).value = line[1]

        fileNum += 2

    workbook.save(filename=writePath)


main()

