#Package imports
from os import listdir
import pandas as pd
import openpyxl as op
import pandas as pd
from datetime import datetime

#Directories of data transfer sources and removal of hidden ds store files
wb_list = []
final_wb_directory = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program Pictet/Upload Template (Positions & CF).xlsx"
final_wb_directory_save = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program Pictet/Upload Final (Positions & CF).xlsx"
path = "/Users/adithyanarayanan/Dropbox/Documents/Personal/Internships/Tristar investment management/Final statement collation/Position statements/Data transfer program Pictet/Excel files"
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

Portfolio_name_list = []
Security_ID_list = []
Type_list = []
Security_Name_list = []
Quantity_list = []
Cost_price = []
Currency_list = []
Exchange_list = []
New_Security_ID_list = []
Date_list = []
Concatenated_lists = [Portfolio_name_list, Type_list, Security_ID_list, Security_Name_list, Quantity_list, Cost_price, Currency_list, Exchange_list, Date_list]
Concatenated_lists2 = [Portfolio_name_list, New_Security_ID_list, Security_Name_list, Quantity_list, Cost_price, Date_list]

for i in wb_list:
    print(i)
    wb_fixinc = pd.read_excel(i)
    print("\nSuccessfully opened file {}".format(i))
    print(wb_fixinc)
    Portfolio_name_list.extend(wb_fixinc[wb_fixinc.columns[0]].values.tolist())
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[43]].values.tolist())
    Type_list.extend(wb_fixinc[wb_fixinc.columns[3]].values.tolist())
    Security_Name_list.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[13]].values.tolist())
    Currency_list.extend(wb_fixinc[wb_fixinc.columns[12]].values.tolist())
    Exchange_list.extend(wb_fixinc[wb_fixinc.columns[17]].values.tolist())
    Date_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())

for i in range(len(Date_list)):
    Date_list[i] = pd.to_datetime(Date_list[i])

# Use values from the list of dates and find the latest value and set that to be the record date for all the transactions
for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        Greatestdate = datetime.strptime(str(Date_list[i]), "%Y-%m-%d %H:%M:%S")
        break

for i in range(len(Date_list)):
    if str(Date_list[i]) != "nan":
        if datetime.strptime(str(Date_list[i]), "%Y-%m-%d %H:%M:%S") > Greatestdate:
            Greatestdate = datetime.strptime(str(Date_list[i]), "%Y-%m-%d %H:%M:%S")

for i in range(len(Date_list)):
    Date_list[i] = Greatestdate.strftime("%d-%m-%Y")

#Addition of Currency tickers for any position without an ISIN:
for i in range(len(Security_Name_list)):
    if str(Security_ID_list[i]) == "nan":
        if "USD" in Currency_list[i]:
            New_Security_ID_list.append("USD Curncy")
        else:
            New_Security_ID_list.append(Currency_list[i] + "USD Curncy")
    else:
        New_Security_ID_list.append(Security_ID_list[i])

for i in range(len(Type_list)):
    if Type_list[i] == "Cash":
        Cost_price[i] = Exchange_list[i]

print(Date_list)

final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

for i in Concatenated_lists2:
    for j in range(8, len(i)+8):
        if i == Portfolio_name_list:
            final_sheet["Q" + str(j)] = i[j-8]
        if i == New_Security_ID_list:
            final_sheet["C" + str(j)] = i[j-8]
        if i == Security_Name_list:
            final_sheet["A" + str(j)] = i[j-8]
        if i == Quantity_list:
            final_sheet["O" + str(j)] = i[j-8]
        if i == Cost_price:
            final_sheet["P" + str(j)] = i[j-8]
        if i == Date_list:
            final_sheet["D" + str(j)] = i[j-8]

final_wb.save(final_wb_directory_save)