import pypdf, openpyxl, os, re, sys
from pdfminer.high_level import extract_pages, extract_text
sys.path.append(r'C:\Users\MOHIBIM\Documents\Coding projects\M-T-Projects\Invoice Reconciliation\helpers.py')
from helpers import *

def main():
    # file to produce the invoice results 
    writePath = r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\Remittance Output.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # User provides the location of the folders where invoices are
    folLoc = input("Provide Month of Remittance in YYMM format: ")
    folLoc = str(checkIfBlank(folLoc,"month"))
    folLoc = r'C:\Users\MOHIBIM\OneDrive - Ventia\\Documents\M&T Finance\Remittance\\' + str(folLoc)
    
    rowNum = 1
    # Walk through the folder and find the pdf files
    for dirpath, direnames, files in os.walk(folLoc):
        for file in files: 
            print(file)
            filename = os.fsdecode(file)

            # Looks through all pdf files
            if filename.endswith(".pdf"):
                newFilePath = os.path.join(dirpath,filename)
                pdfText = extract_text(newFilePath)

                # Looks for an eight digit number beginnign with 93 which all invoice numbers do 
                # and transfers that into a list to be sent to the excel
                invoiceNums = [(pdfText[x.end()-9:x.end()]).strip() for x in re.finditer(r"93([0-9a-zA-Z]){6}", pdfText)]
                print(invoiceNums)

            else:
                print(f'{filename} is not a pdf')
            
             # Print to the excel sheet the invoice number and values per row, this way we can xlookup the output without changing the present sheet
            
            for num in invoiceNums:
                sheet.cell(row= rowNum,column=1).value = num
                rowNum += 1

        

                             
    workbook.save(filename=writePath)

main()