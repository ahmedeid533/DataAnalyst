# import required modules
import pandas as pd
import numpy as np
import time
  
# read data from file to delete dublecate
dfTable = pd.read_excel("Company.xlsx")
DfTable_len = dfTable.shape[0]
#get first custumer ID
custumer = dfTable['Prescription Number'][0]
# first custumer ID position
dfCustID = 0
itrInner = 0
while (dfCustID < DfTable_len - 2):
    if ((dfCustID % 1000) == 0):
        print(str("%.3f" % ((dfCustID / DfTable_len) * 100)) +"%")
    # get View small Table for each cutumer
    dfCust = dfTable.loc[dfTable['Prescription Number'] == custumer]
    # number of rows for each custumer Table
    dfCustRows = dfCust.shape[0]
    
    if (dfCustRows >= 2):
        i = 0
        while i < dfCustRows-1:
            # if the operation done twice remove one and increase quantity
            j = i + 1
            if ((dfCustID + i) not in dfTable.index) :
                i += 1
                continue

            while  j <= dfCustRows-1:
                if ((dfCustID + j) not in dfTable.index):
                    j += 1
                    continue

                check_1 = dfCust['Item Dispensed Number'][dfCustID + i] != dfCust['Item Dispensed Number'][dfCustID + j]
                check_2 = dfCust['Pharmacy Branch Code'][dfCustID + i] != dfCust['Pharmacy Branch Code'][dfCustID + j]
                check_3 = dfCust['Date of Dispensing'][dfCustID + i] != dfCust['Date of Dispensing'][dfCustID + j]
                if (check_1 | check_2 | check_3):
                    j += 1
                    continue

                boolean_t = dfTable['Transaction No'][dfCustID + i] !=  dfTable['Transaction No'][dfCustID + j]
                if (boolean_t):
                    dfTable.at[dfCustID + i, 'dublicate'] = "yes"
                dfTable.at[dfCustID + i, 'Quantity'] = dfTable['Quantity'][dfCustID + i] + dfTable['Quantity'][dfCustID + j]
                dfTable.at[dfCustID + i, 'Net Amount'] = dfTable['Net Amount'][dfCustID + i] + dfTable['Net Amount'][dfCustID + j]
                dfTable = dfTable.drop(index=(dfCustID + j))
                j += 1
            i += 1
    dfCustID += dfCustRows
    if dfCustID >= DfTable_len-1:
        break
    custumer = dfTable['Prescription Number'][dfCustID]
print("100%")
dfTable.to_excel('Company_remove_dublicate.xlsx') 
