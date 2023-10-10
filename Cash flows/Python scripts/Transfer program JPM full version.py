from os import listdir
import pandas as pd
import openpyxl as op

wb_list = []
final_wb_directory = input("Please enter the directory of the file you would like to use as the template:")
final_wb_directory_save = input("Please enter the directory with the file name you would like to save changes to:")
path = input("Please enter the directory of the folder you would like to add:")
month = input("Please enter the month you would like to extract the data for (MM):")
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
Description_list = []
Type_list = []
Amount_list =[]
Amountnonbase_list = []
Quantity_list = []
Cost_price = []
Currency_list = []
Currency_conv_list = []
Record_date_list = []
Settlement_date_list = []
Security_ID_list = []

Portfolio_name_list_final = []
Description_list_final = []
Type_list_final = []
Amount_list_final = []
Amountnonbase_list_final = []
Quantity_list_final = []
Cost_price_final = []
Currency_list_final = []
Currency_conv_list_final = []
Record_date_list_final = []
Settlement_date_list_final = []
Security_ID_list_final = []

Portfolio_name_list_final2 = []
Description_list_final2 = []
Type_list_final2 = []
Amount_list_final2 = []
Amountnonbase_list_final2 = []
Quantity_list_final2 = []
Cost_price_final2 = []
Currency_list_final2 = []
Currency_conv_list_final2 = []
Record_date_list_final2 = []
Settlement_date_list_final2 = []
Security_ID_list_final2 = []

New_Quantity_list = []
New_Type_list = []
New_Price_list = []
New_Security_ID_list = []

Concatenated_lists = [Portfolio_name_list, Description_list, Type_list, Amount_list, Amountnonbase_list, Quantity_list, Cost_price, Currency_list, Currency_conv_list, Record_date_list, Settlement_date_list, Security_ID_list]
Concatenated_lists_final = [Portfolio_name_list_final, Description_list_final, Type_list_final, Amount_list_final, Amountnonbase_list_final, Quantity_list_final, Cost_price_final, Currency_list_final, Currency_conv_list_final, Record_date_list_final, Settlement_date_list_final, Security_ID_list_final]
Concatenated_lists_final2 = [Portfolio_name_list_final2, Description_list_final2, Type_list_final2, Amount_list_final2, Amountnonbase_list_final2, Quantity_list_final2, Cost_price_final2, Currency_list_final2, Currency_conv_list_final2, Record_date_list_final2, Settlement_date_list_final2, Security_ID_list_final2]
Concatenated_lists_final3 = [Portfolio_name_list_final2, Description_list_final2, New_Type_list, New_Quantity_list, New_Price_list, Record_date_list_final2, Settlement_date_list_final2, New_Security_ID_list]
Concatenated_lists_final4 = [Portfolio_name_list_final2, Description_list_final2, New_Type_list, Amount_list_final2, Amountnonbase_list_final2, New_Quantity_list, New_Price_list, Currency_list_final2, Currency_conv_list_final2, Record_date_list_final2, Settlement_date_list_final2, New_Security_ID_list]

def transferdata(Portfolio_name_list, Description_list, Type_list, Amount_list, Amountnonbase_list, Quantity_list, Cost_price, Currency_list, Currency_conv_list, Record_date_list, Settlement_date_list, Security_ID_list, wb_fixinc):
    Portfolio_name_list.extend(wb_fixinc[wb_fixinc.columns[0]].values.tolist())
    Description_list.extend(wb_fixinc[wb_fixinc.columns[1]].values.tolist())
    Type_list.extend(wb_fixinc[wb_fixinc.columns[2]].values.tolist())
    Amount_list.extend(wb_fixinc[wb_fixinc.columns[3]].values.tolist())
    Amountnonbase_list.extend(wb_fixinc[wb_fixinc.columns[12]].values.tolist())
    Quantity_list.extend(wb_fixinc[wb_fixinc.columns[5]].values.tolist())
    Cost_price.extend(wb_fixinc[wb_fixinc.columns[6]].values.tolist())
    Currency_list.extend(wb_fixinc[wb_fixinc.columns[11]].values.tolist())
    Currency_conv_list.extend(wb_fixinc[wb_fixinc.columns[20]].values.tolist())
    Security_ID_list.extend(wb_fixinc[wb_fixinc.columns[7]].values.tolist())
    Record_date_list.extend(wb_fixinc[wb_fixinc.columns[8]].values.tolist())
    Settlement_date_list.extend(wb_fixinc[wb_fixinc.columns[10]].values.tolist())

for i in wb_list:
    print(i)
    wb_pd = pd.ExcelFile(i)
    print("Opening workbook {}".format(str(i)))
    if "Fixed Income & Cash" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Fixed Income & Cash", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Fixed Income & Cash"))
        print(wb_fixinc)
        transferdata(Portfolio_name_list, Description_list, Type_list, Amount_list, Amountnonbase_list, Quantity_list, Cost_price, Currency_list, Currency_conv_list, Record_date_list, Settlement_date_list, Security_ID_list, wb_fixinc)
    if "Equity" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Equity", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Equity"))
        print(wb_fixinc)
        transferdata(Portfolio_name_list, Description_list, Type_list, Amount_list, Amountnonbase_list, Quantity_list, Cost_price, Currency_list, Currency_conv_list, Record_date_list, Settlement_date_list, Security_ID_list, wb_fixinc)
    if "Alternative Assets" in wb_pd.sheet_names:
        wb_fixinc = pd.read_excel(i, sheet_name="Alternative Assets", skiprows=1)
        print("\nSuccessfully opened sheet {}".format("Alternative Assets"))
        print(wb_fixinc)
        transferdata(Portfolio_name_list, Description_list, Type_list, Amount_list, Amountnonbase_list, Quantity_list, Cost_price, Currency_list, Currency_conv_list, Record_date_list, Settlement_date_list, Security_ID_list, wb_fixinc)

for m in range(len(Record_date_list)):
    if Record_date_list[m][5] == month[0] and Record_date_list[m][6] == month[1]:
        Portfolio_name_list_final.append(Portfolio_name_list[m])
        Description_list_final.append(Description_list[m])
        Type_list_final.append(Type_list[m])
        Amount_list_final.append(Amount_list[m])
        Amountnonbase_list_final.append(Amountnonbase_list[m])
        Quantity_list_final.append(Quantity_list[m])
        Cost_price_final.append(Cost_price[m])
        Currency_list_final.append(Currency_list[m])
        Currency_conv_list_final.append(Currency_conv_list[m])
        Record_date_list_final.append(Record_date_list[m])
        Settlement_date_list_final.append(Settlement_date_list[m])
        Security_ID_list_final.append(Security_ID_list[m])


for m in range(len(Description_list_final)):
    if "SAS/" not in Description_list_final[m] and "PUS/" not in Description_list_final[m] and "REM/" not in Description_list_final[m] and "DIV/" not in Description_list_final[m] and "CPS/" not in Description_list_final[m] and "SUB/" not in Description_list_final[m]:
        Portfolio_name_list_final2.append(Portfolio_name_list_final[m])
        Description_list_final2.append(Description_list_final[m])
        Type_list_final2.append(Type_list_final[m])
        Amount_list_final2.append(Amount_list_final[m])
        Amountnonbase_list_final2.append(Amountnonbase_list_final[m])
        Quantity_list_final2.append(Quantity_list_final[m])
        Cost_price_final2.append(Cost_price_final[m])
        Currency_list_final2.append(Currency_list_final[m])
        Currency_conv_list_final2.append(Currency_conv_list_final[m])
        Record_date_list_final2.append(Record_date_list_final[m])
        Settlement_date_list_final2.append(Settlement_date_list_final[m])
        Security_ID_list_final2.append(Security_ID_list_final[m])

for i in range(len(Type_list_final2)):
    if Type_list_final2[i] == "Redemption" or Type_list_final2[i] == "Sale":
        New_Type_list.append("Sell")
    elif Type_list_final2[i] == "Purchase" or Type_list_final2[i] == "Subscription":
        New_Type_list.append("Buy")
    elif "CREDIT AS OF" in Description_list_final2[i] or "ADJUST. CREDIT INT." in Description_list_final2[i]:
        New_Type_list.append("Buy")
    elif "Expenses payment" in Description_list_final2[i] or "Audit Fees" in Description_list_final2[i]:
        New_Type_list.append("Management_Fee")
    elif "OUTGOING REMITTANCE" in Description_list_final2[i]:
        New_Type_list.append("Withdrawal")
    elif Quantity_list_final2[i] < 0:
        New_Type_list.append("Sell")
    elif Quantity_list_final2[i] > 0:
        New_Type_list.append("Buy")
    else:
        New_Type_list.append("")

for i in range(len(Cost_price_final2)):
    if "CREDIT AS OF" in Description_list_final2[i] or "ADJUST. CREDIT INT." in Description_list_final2[i]:
        New_Price_list.append(0)
    elif "Expenses payment" in Description_list_final2[i] or "Audit Fees" in Description_list_final2[i] or "OUTGOING REMITTANCE" in Description_list_final2[i]:
        New_Price_list.append(1)
    else:
        New_Price_list.append(Cost_price_final2[i])

for i in range(len(Quantity_list_final2)):
    if "CREDIT AS OF" in Description_list_final2[i] or "ADJUST. CREDIT INT." in Description_list_final2[i] or "Expenses payment" in Description_list_final2[i] or "Audit Fees" in Description_list_final2[i] or "OUTGOING REMITTANCE" in Description_list_final2[i]:
        New_Quantity_list.append(Amount_list_final2[i])
    else:
        New_Quantity_list.append(Quantity_list_final2[i])

for i in range(len(Security_ID_list_final2)):
    if pd.isna(Security_ID_list_final2[i]):
        if Currency_list_final2[i] == "USD":
            New_Security_ID_list.append(Currency_list_final2[i] + " Curncy")
        else:
            New_Security_ID_list.append(Currency_list_final2[i] + "USD Curncy")
    else:
        New_Security_ID_list.append(Security_ID_list_final2[i])

for i in range(len(Description_list_final2)):
    if " SPOT " in Description_list_final2[i] and Amount_list_final2[i] < 0:
        New_Type_list[i] = "Sell"
    elif " SPOT " in Description_list_final2[i] and Amount_list_final2[i] > 0:
        New_Type_list[i] = "Buy"

for i in range(len(Description_list_final2)):
    if " SPOT " in Description_list_final2[i]:
        New_Quantity_list[i] = Amountnonbase_list_final2[i]
        New_Price_list[i] = Currency_conv_list_final2[i]

for i in range(len(New_Quantity_list)):
    if New_Quantity_list[i]<0:
        New_Quantity_list[i] = New_Quantity_list[i] * -1

m = 0
while m < len(Description_list_final2):
    if " SPOT " in Description_list_final2[m] and Currency_list_final2[m] == "USD":
        print(Description_list_final2[m], Currency_list_final2[m])
        for i in Concatenated_lists_final4:
            del i[m]
    m=m+1


for i in Concatenated_lists_final3:
    print(str(i) + "\n")

final_wb = op.load_workbook(final_wb_directory)
final_sheet = final_wb.get_sheet_by_name("template")

for i in Concatenated_lists_final3:
    for j in range(8, len(i)+8):
        if i == Portfolio_name_list_final2:
            final_sheet["Q" + str(j)] = i[j-8]
        if i == Description_list_final2:
            final_sheet["A" + str(j)] = i[j-8]
        if i == New_Type_list:
            final_sheet["B" + str(j)] = i[j-8]
        if i == New_Quantity_list:
            final_sheet["O" + str(j)] = i[j-8]
        if i == New_Price_list:
            final_sheet["P" + str(j)] = i[j-8]
        if i == New_Security_ID_list:
            final_sheet["C" + str(j)] = i[j-8]
        if i == Record_date_list_final2:
            final_sheet["D" + str(j)] = i[j-8]
        if i == Settlement_date_list_final2:
            final_sheet["E" + str(j)] = i[j-8]

final_wb.save(final_wb_directory_save)
