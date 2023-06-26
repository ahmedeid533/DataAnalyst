# import required modules
import pandas as pd
import numpy as np
  
# read data Reference from file Reference.xlsx (main file)
dfTable_Reference = pd.read_excel('Reference.xlsx')
dfTable_Reference_len = dfTable_Reference.shape[0]

# read data Reference from file Company_remove_dublicate.xlsx
dfTable_Company = pd.read_excel('Company_remove_dublicate.xlsx')
dfTable_Company_len = dfTable_Company.shape[0]

# DFTABLE For cust in Reference not in Company
dfCustInReference = pd.DataFrame()

# DFTABLE For cust in Company not in Reference 
dfCustInCompany = pd.DataFrame()

# DFTABLE iam a row in Reference
dfRowInReference = pd.DataFrame()

# DFTABLE iam a row in Company
dfRowInCompany = pd.DataFrame()

# empty rows for ease porpose
emline_Reference = pd.Series(["Nan"])
emline_Company = pd.Series(["Nan"])

#get first custumer ID
custumerReference = dfTable_Reference['Prescription Number'][0]
custumerCompany = dfTable_Company['Prescription Number'][0]

# first custumer ID position
dfCustIdReference = 0
dfCustIdCompany = 0
while (dfCustIdReference < dfTable_Reference_len - 1):
    # get the ID
    custumerReference = dfTable_Reference['Prescription Number'][dfCustIdReference]
    custumerCompany = dfTable_Company['Prescription Number'][dfCustIdCompany]

    # percentage of completation
    if ((dfCustIdReference % 1000) == 0):
        print(str("%.1f" % ((dfCustIdReference / dfTable_Reference_len) * 100)) +"%")

    # get View small Table for each cutumer
    dfCust_Reference = dfTable_Reference.loc[dfTable_Reference['Prescription Number'] == custumerReference]
    dfCust_Company = dfTable_Company.loc[dfTable_Company['Prescription Number'] == custumerCompany]

    # get a copy from the view
    dfCust_Reference_copy = dfCust_Reference.reset_index(drop=True).copy()
    dfCust_Company_copy = dfCust_Company.reset_index(drop=True).copy()

    # number of rows for each custumer Table
    dfCustRows_Reference = dfCust_Reference.shape[0]
    dfCustRows_Company = dfCust_Company.shape[0]
    dfCustRows_Company_C = dfCustRows_Company
    
    if (custumerReference == custumerCompany):
        iReference = 0
        while iReference < dfCustRows_Reference:
            jCompany = 0
            while jCompany < dfCustRows_Company_C:
                boolean_Name = (dfCust_Reference_copy['ItemName'][iReference] == dfCust_Company_copy['ItemName'][jCompany])
                boolean_Quantity = (dfCust_Reference_copy['Qty'][iReference] >= dfCust_Company_copy['Quantity'][jCompany])
                boolean_Date = (dfCust_Reference_copy['Dispense Date'][iReference] == dfCust_Company_copy['Date of Dispensing'][jCompany])
                if (boolean_Name & boolean_Quantity & boolean_Date):
                    dfCust_Company_copy = dfCust_Company_copy.drop([jCompany]).reset_index(drop=True)
                    dfCustRows_Company_C -= 1
                    dfCust_Reference_copy = dfCust_Reference_copy.drop([iReference])
                    break
                jCompany += 1
            iReference += 1
        
        # for beauty poropse
        tempRowsCompany = dfCust_Company_copy.shape[0]
        tempRowsReference = dfCust_Reference_copy.shape[0]
        while tempRowsReference < tempRowsCompany :
            empReference = emline_Reference.copy()
            dfCust_Reference_copy = pd.concat([dfCust_Reference_copy, empReference])
            tempRowsReference += 1

        while tempRowsCompany < tempRowsReference :
            empCompany = emline_Company.copy()
            dfCust_Company_copy = pd.concat([dfCust_Company_copy, empCompany])
            tempRowsCompany += 1

        dfRowInReference = pd.concat([dfRowInReference, dfCust_Reference_copy])      
        dfRowInCompany = pd.concat([dfRowInCompany, dfCust_Company_copy])
        dfCustIdReference += dfCustRows_Reference
        dfCustIdCompany += dfCustRows_Company
    elif (custumerReference > custumerCompany):
        dfCustInCompany = pd.concat([dfCustInCompany, dfCust_Company])
        dfCustIdCompany += dfCustRows_Company
    else :
        dfCustInReference = pd.concat([dfCustInReference, dfCust_Reference])
        dfCustIdReference += dfCustRows_Reference


print("100%")


dfCustInReference.to_excel('Reference_Cust.xlsx') 
dfCustInCompany.to_excel('Company_Cust.xlsx')
dfRowInReference.to_excel('changesInReference.xlsx')
dfRowInCompany.to_excel('changesInCompany.xlsx')