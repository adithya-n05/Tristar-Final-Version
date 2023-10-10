#Package imports
from os import listdir
import pandas as pd
import openpyxl as op
from datetime import datetime


wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program BOS/Upload Template (Positions & CF).xlsx" 
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program BOS/Upload Final (Positions & CF).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program BOS/Excel files"
path = path + "/"
dir_list = listdir(path)
for i in range(len(dir_list)):
    dir_list[i] = path + dir_list[i]
ds_store_path = path + ".DS_Store"
if ds_store_path in dir_list:
    dir_list.remove(ds_store_path)
wb_list = dir_list
print("\nThe workbooks you have added are:")
print(wb_list)

#Defining lists to copy data into
Portfolio_name_list = []
Security_ID_list = []
Security_Name_list = []
Quantity_list = []
Cost_price = []
Market_Value_Orig_Currency_list = []
Market_Value_USD_list = []
Currencylist = []
Concatenated_lists = [Portfolio_name_list, Security_ID_list, Security_Name_list, Quantity_list, Cost_price, Market_Value_Orig_Currency_list, Market_Value_USD_list, Currencylist]

#Defining final lists to copy data into
Portfolio_name_list_final = []
Security_ID_list_final = []
New_security_ID_list_final = [] #This list will fill in the gaps from the security ID list for cash accounts
Purchase_date_list_final = []
Security_Name_list_final = []
Quantity_list_final = []
Cost_price_final = []
Market_Value_Orig_Currency_list_final = []
Market_Value_USD_list_final = []
Currencylist_final = []
Concatenated_lists_final = [Portfolio_name_list_final, New_security_ID_list_final, Purchase_date_list_final, Security_Name_list_final, Quantity_list_final, Cost_price_final, Market_Value_Orig_Currency_list_final, Market_Value_Orig_Currency_list_final, Currencylist_final]

#Reading excel file one
wb_fixinc = pd.read_excel(wb_list[0])
print(wb_fixinc)

GreatestDate = wb_fixinc.iloc[1,1]
GreatestDate = GreatestDate[:-4]
print(GreatestDate)

for i in wb_list:
    #Reading excel files
    wb_fixinc = pd.read_excel(i)
    print("\nSuccessfully opened file {}".format(i))
    print(wb_fixinc) 
    if datetime.strptime(str(wb_fixinc.iloc[1,1][:-4]), "%d-%m-%Y %H:%M:%S") > datetime.strptime(GreatestDate, "%d-%m-%Y %H:%M:%S"):
        GreatestDate = wb_fixinc.iloc[1,1][:-3]

    #Transferring data from columns to lists within python
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[3]].values.tolist())
    Security_Name_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Market_Value_Orig_Currency_list.extend(wb_fixinc[wb_fixinc.columns[10]].values.tolist())
    Market_Value_USD_list.extend(wb_fixinc[wb_fixinc.columns[11]].values.tolist())
    Currencylist.extend(wb_fixinc[wb_fixinc.columns[4]].values.tolist())

    #Removing empty spaces that come from first 11 rows of the files
    del Security_ID_list[0:11]
    del Security_Name_list[0:11]
    del Quantity_list[0:11]
    del Cost_price[0:11]
    del Market_Value_Orig_Currency_list[0:11]
    del Market_Value_USD_list[0:11]
    del Currencylist[0:11]

    #Pushing data to final list and resetting temporary list that carries data from the excel files to the final file
    #This is done as if a new file's data is appended to the original list, the location of the 11 rows in front cannot be found
    Security_ID_list_final.extend(Security_ID_list) 
    Security_Name_list_final.extend(Security_Name_list)
    Quantity_list_final.extend(Quantity_list)
    Cost_price_final.extend(Cost_price)
    Market_Value_Orig_Currency_list_final.extend(Market_Value_Orig_Currency_list)
    Market_Value_USD_list_final.extend(Market_Value_USD_list)
    Currencylist_final.extend(Currencylist)
    Portfolio_name_list = []
    Security_ID_list = []
    Security_Name_list = []
    Quantity_list = []
    Cost_price = []
    Market_Value_Orig_Currency_list = []
    Market_Value_USD_list = []
    Currencylist = []

#Setting all of the Portfolio name list to be the ID of the security
for i in Security_ID_list_final:
    Portfolio_name_list_final.append(wb_fixinc.iloc[3,1])

#Setting all of the date list to be the single date on the sheet
for i in Security_ID_list_final:
    Purchase_date_list_final.append(GreatestDate[:-9])

#Printing of all lists for debugging
for i in Concatenated_lists_final:
    print(i)

#Checking and changing the values of the Price and Quantity values based on whether "Current account" is found in the name of the security
for i in range(len(Security_Name_list_final)):
    if "Current Account" in Security_Name_list_final[i]:
        if str(Market_Value_USD_list_final[i]) != " ":
            Quantity_list_final[i] = Market_Value_Orig_Currency_list_final[i]
            Cost_price_final[i] = float(str(Market_Value_Orig_Currency_list_final[i]).replace(",", ""))/float(str(Market_Value_USD_list_final[i]).replace(",", ""))
        else:
            Quantity_list_final[i] = 0
            Cost_price_final[i] = Market_Value_Orig_Currency_list_final[i]


#Addition of Currency tickers for cash accounts:
for i in range(len(Security_Name_list_final)):
    if "Current Account" in Security_Name_list_final[i] or "External Securities" in Security_Name_list_final[i]:
        if "USD" in Currencylist_final[i]:
            New_security_ID_list_final.append("USD Curncy")
        else:
            New_security_ID_list_final.append(Currencylist_final [i] + "USD Curncy")
    else:
        New_security_ID_list_final.append(Security_ID_list_final[i])

#Opening template workbook and sheet of name "template" within it
final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

#Transferring data from lists into final template file and saving using directory given to save to
for i in Concatenated_lists_final:
    for j in range(8, len(i)+8):
        if i == Portfolio_name_list_final:
            final_sheet["Q" + str(j)] = i[j-8]
        if i == New_security_ID_list_final:
            final_sheet["C" + str(j)] = i[j-8]
        if i == Purchase_date_list_final:
            final_sheet["D" + str(j)] = i[j-8]
        if i == Security_Name_list_final:
            final_sheet["A" + str(j)] = i[j-8]
        if i == Quantity_list_final:
            final_sheet["O" + str(j)] = i[j-8]
        if i == Cost_price_final:
            final_sheet["P" + str(j)] = i[j-8]

final_wb.save(final_wb_directory_save)