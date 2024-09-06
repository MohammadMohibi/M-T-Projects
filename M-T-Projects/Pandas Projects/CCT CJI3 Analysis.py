import pandas as pd

# Read the CCT File where we will be sorting data to
cct = pd.read_excel(r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\CCT Cji3 costs from aug 2022 to aug 2024.xlsx')

col = cct.columns

# Allocate the priamry groupings by column - Declutter data
cctPrime = cct[['Posting Date', 'WBS Name', 'Vendor Name', 'Purchase order text', 'Cost element descr.', 'Purchasing Document']]

for index, row in cctPrime.iterrows():
    print(index, row)
