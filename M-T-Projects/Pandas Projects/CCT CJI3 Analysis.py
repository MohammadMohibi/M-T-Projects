import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'

# Read the CCT File where we will be sorting data to
readPath = r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\CCT Cji3 costs from aug 2022 to aug 2024.xlsx'
writePath = r'C:\Users\MOHIBIM\OneDrive - Ventia\Documents\M&T Finance\output.xlsx'
# dLambda = lambda x: pd.datetime.strptime(x, '%Y-%m')
cct = pd.read_excel(readPath)

col = cct.columns

# Allocate the priamry groupings by column - Declutter data
cctPrime = cct[['Posting Date', 'WBS Name', 'Vendor Name', 'Value TranCurr', 'Purchase order text', 'Cost element descr.', 'Purchasing Document']]


# Change the date posted to the peiod for a better aggregate sum in that period
cctPrime['Period'] = cctPrime.iloc[:,0].dt.to_period('M')
cols = list(cctPrime.columns)
cctPrime = cctPrime[[cols[-1]]+ cols[1:-1]]

# Begin grouping by the necessary variables WBS -> Vendor -> aggregate sum based on month
#wbsGroup = cctPrime.groupby(['WBS Name', 'Vendor Name', 'Period']).agg('Value TranCurr', sum).unstack('Period', fill_value=0)
POTextGroup = pd.DataFrame()
wbsGroup = cctPrime.groupby(['WBS Name', 'Vendor Name', 'Period'])['Value TranCurr'].sum().unstack('Period', fill_value='')
POTextGroup['PO'] = cctPrime.groupby(['WBS Name', 'Vendor Name', 'Period'])['Purchase order text'].apply('-'.join)
POTextGroup['Tran'] = cctPrime.groupby(['WBS Name', 'Vendor Name', 'Period'])['Value TranCurr'].apply(str)
#wbsGroup = pd.DataFrame(wbsGroup).pivot(columns='Period', values='Value TranCurr')

# Combine the two columns so they can be unstacked together and isolate it 
POTextGroup['PO/Tran'] = POTextGroup[['PO', 'Tran']].apply(lambda row: '_'.join(row.values.astype(str)), axis = 1)
POTextGroup = POTextGroup[['PO/Tran']].unstack('Period', fill_value='')

# Add the indexation so excel output looks better
wbsGroup = wbsGroup.rename_axis(columns=None).reset_index()
POTextGroup = POTextGroup.reset_index()

print(wbsGroup)
print(POTextGroup)

# Transfer to excel file in the same location
writer = pd.ExcelWriter(writePath, engine='openpyxl')
wbsGroup.to_excel(writer, index = False, sheet_name='Sum Only')
POTextGroup.to_excel(writer, index = True, sheet_name='PO Text')

# Close out writer 
writer.close()