import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'

# Read the CCT File where we will be sorting data to
# dLambda = lambda x: pd.datetime.strptime(x, '%Y-%m')
cct = pd.read_excel(r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\CCT Cji3 costs from aug 2022 to aug 2024.xlsx')

col = cct.columns

# Allocate the priamry groupings by column - Declutter data
cctPrime = cct[['Posting Date', 'WBS Name', 'Vendor Name', 'Value TranCurr', 'Purchase order text', 'Cost element descr.', 'Purchasing Document']]


# Change the date posted to the peiod for a better aggregate sum in that period
cctPrime['Period'] = cctPrime.iloc[:,0].dt.to_period('M')
cols = list(cctPrime.columns)
cctPrime = cctPrime[[cols[-1]]+ cols[1:-1]]

# Begin grouping by the necessary variables WBS -> Vendor -> aggregate sum based on month
wbsGroup = pd.DataFrame(cctPrime.groupby(['WBS Name', 'Vendor Name']))

wbsGroup.to_excel(r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\output.xlsx', index = False)
print(wbsGroup.first())


'''for index, row in cctPrime.iterrows():
    print(index, row)'''
