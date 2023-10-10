#Package imports
from os import listdir
import pandas as pd
import openpyxl as op
from datetime import datetime

#Directories of data transfer sources and removal of hidden .DS_Store files
wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program UBS/Upload Template (Positions & CF).xlsx"
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program UBS/Upload Final (Positions & CF).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program UBS/CSV files"
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
Column_1_list = []
Portfolio_name_list = []
Security_ID_list = []
New_security_ID_list = [] #This list will fill in the gaps from the security ID list for cash accounts
Security_Name_list_col1 = []
Security_Name_list_col2 = []
Security_Name_list_finalcol = []
Quantity_list = []
Cost_price = []
New_price_list = []
New_price_list2 = []
Currencylist = []
Market_value_list = []
Date_list = []
Concatenated_lists_prior = [Column_1_list, Portfolio_name_list, Security_ID_list, Security_Name_list_col1, Security_Name_list_col2, Quantity_list, Cost_price, Currencylist, Market_value_list, Date_list]
Concatenated_lists = [Column_1_list, Portfolio_name_list, Security_ID_list, Security_Name_list_finalcol, Quantity_list, Cost_price, Currencylist, Market_value_list,  Date_list]
Concatenated_lists_final = [Column_1_list, Portfolio_name_list, New_security_ID_list, Security_Name_list_finalcol, Quantity_list, New_price_list2, Date_list]

for i in wb_list:
    print(i)
    wb_fixinc = pd.read_csv(i, sep=';')
    print("\nSuccessfully opened file {}".format(i))
    print(wb_fixinc)
    Column_1_list.extend(wb_fixinc[wb_fixinc.columns[0]].values.tolist())
    Portfolio_name_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[8]].values.tolist())
    Security_Name_list_col1.extend(wb_fixinc[wb_fixinc.columns[13]].values.tolist())
    Security_Name_list_col2.extend(wb_fixinc[wb_fixinc.columns[14]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[9]].values.tolist())
    Currencylist.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Market_value_list.extend(wb_fixinc[wb_fixinc.columns[23]].values.tolist())
    Date_list.extend(wb_fixinc[wb_fixinc.columns[21]].values.tolist())
    n=len(Column_1_list)-1

    #Removing empty rows
    while n>0:
        if Column_1_list[n] == "Detailed positions: Liquidity - Accounts from":
            print(Column_1_list[n], n)
            for i in Concatenated_lists_prior:
                del i[n:len(i)]
            break
        n=n-1

#Merging two security description columns together into one
for i in range(len(Security_Name_list_col1)):
    Security_Name_list_finalcol.append(Security_Name_list_col1[i] + Security_Name_list_col2[i])

#Removal of empty header rows using key "portfolio" as a landmark
m=0
while m < len(Portfolio_name_list):
    if Portfolio_name_list[m] =="Portfolio":
        for k in Concatenated_lists:
            k.pop(m-1)
            k.pop(m-1)
    m=m+1

#Remove apostrophe for thousands delimiter
for k in range(len(Quantity_list)):
    Quantity_list[k]=Quantity_list[k].replace("'","")

for k in range(len(Cost_price)):
    replace_value = Cost_price[k].replace("'","")
    Cost_price[k]=replace_value

for k in range(len(Market_value_list)):
    replace_value = Market_value_list[k].replace("'","")
    Market_value_list[k]=replace_value

#Addition of Currency tickers for cash accounts:
for i in range(len(Security_Name_list_finalcol)):
    if "Current Account" in Security_Name_list_finalcol[i]:
        if "USD" in Currencylist[i]:
            New_security_ID_list.append("USD Curncy")
        else:
            New_security_ID_list.append(Currencylist[i] + "USD Curncy")
    else:
        New_security_ID_list.append(Security_ID_list[i])

# If the price contains a %, it will be removed
for i in range(len(Cost_price)):
    if "%" in str(Cost_price[i]):
        New_price_list.append(str(Cost_price[i]).replace("%", ""))
    else:
        New_price_list.append(str(Cost_price[i]))

# Take the market value column divided by the Number/Amt for the price column if the description contains current account. If the denominator is 0, the value is automatically set to 0.
for i in range(len(Security_Name_list_finalcol)):
    if "Current Account" in Security_Name_list_finalcol[i]:
        if float(Quantity_list[i]) != 0:
            New_price_list2.append(float(Market_value_list[i])/float(Quantity_list[i]))
        else:
            New_price_list2.append(0)
    else:
        New_price_list2.append(New_price_list[i])

# Use values from the list of dates and find the latest value and set that to be the record date for all the transactions
for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        Greatestdate = datetime.strptime(str(Date_list[i]), "%d.%m.%Y")
        break

for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        if datetime.strptime(str(Date_list[i]), "%d.%m.%Y") > Greatestdate:
            Greatestdate = datetime.strptime(str(Date_list[i]), "%d.%m.%Y")

for i in range(len(Date_list)):
    Date_list[i] = Greatestdate.strftime("%d.%m.%Y")

#Opening template workbook and sheet of name "template" within it
final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

#Transferring data from lists into final template file and saving using directory given to save to
for i in Concatenated_lists_final:
    for j in range(8, len(i)+8):
        if i == Portfolio_name_list:
            final_sheet["Q" + str(j)] = i[j-8]
        if i == New_security_ID_list:
            final_sheet["C" + str(j)] = i[j-8]
        if i == Security_Name_list_finalcol:
            final_sheet["A" + str(j)] = i[j-8]
        if i == Quantity_list:
            final_sheet["O" + str(j)] = i[j-8]
        if i == New_price_list2:
            final_sheet["P" + str(j)] = i[j-8]
        if i == Date_list:
            final_sheet["D" + str(j)] = i[j-8]

final_wb.save(final_wb_directory_save)