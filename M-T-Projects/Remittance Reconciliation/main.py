import pypdf, openpyxl, os, re, sys
sys.path.append('C:\Users\MOHIBIM\Documents\Coding projects\M-T-Projects\Invoice Reconciliation')
from helpers import *

def main():
    # file to produce the invoice results 
    writePath = r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\Remittance Output.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # User provides the location of the folders where invoices are
    folLoc = input("Provide Month of Remittance in YYMM format: ")
    folLoc = str(checkIfBlank(folLoc,"month"))
    folLoc = r"C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\Remittance\" str(folLoc)
    
    exit()
main()