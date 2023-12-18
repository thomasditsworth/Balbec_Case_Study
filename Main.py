import pandas as pd

def question1(q1_filepath, excel_file_path):
    column_6th = q1_filepath.iloc[:, 5]  # Property ID
    column_7th = q1_filepath.iloc[:, 6]  # Asset Type

    right_table_dict = {}

    for (value_6th, value_7th) in zip(column_6th, column_7th):
        if not pd.isnull(value_6th) and not pd.isnull(value_7th) and (value_6th != 'Property ID'):
            right_table_dict[value_6th] = value_7th

    column_2nd = q1_filepath.iloc[:, 1] # left Property Id
    string_to_avoid = "Use a formula to pick out the asset type for the properties on the left from the table on the right."

    for index, value_2nd in enumerate(column_2nd):
        if not pd.isnull(value_2nd) and (value_2nd != string_to_avoid):
            temp = right_table_dict.get(value_2nd, "Default")
            q1_filepath.iat[int(index), 2] = temp

    # Use ExcelWriter to append to the original Excel file
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 1'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q1_filepath.to_excel(writer, sheet_name='Question 1', index=False, header=False)

def question2(q2_filepath, excel_file_path):

    # incorrect way to do it
    for x in range(6,13):
        val = q2_filepath.iat[x, 1]
        for y in range(x, x + 10):
            temp_val = q2_filepath.iat[y, 4]
            if val == temp_val:
                q2_filepath.iat[x, 2] = q2_filepath.iat[y,5]

    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 2'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q2_filepath.to_excel(writer, sheet_name='Question 2', index=False, header=False)

def question2_real(q2_filepath, excel_file_path):
    column_5th = q2_filepath.iloc[:, 4]  # Property ID
    column_6th = q2_filepath.iloc[:, 5]  # Asset Type

    right_table_dict = {}

    for (value_5th, value_6th) in zip(column_5th, column_6th):
        if not pd.isnull(value_5th) and not pd.isnull(value_6th) and (value_6th != 'ID'):
            right_table_dict[value_5th] = value_6th

    column_2nd = q2_filepath.iloc[:, 1] # left Property Id
    string_to_avoid = "What is wrong with the formula in the table to the left? (2 reasons)"

    for index, value_2nd in enumerate(column_2nd):
        if not pd.isnull(value_2nd) and (value_2nd != string_to_avoid):
            temp = right_table_dict.get(value_2nd, "Default")
            q2_filepath.iat[int(index), 2] = temp

    # Use ExcelWriter to append to the original Excel file
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 2'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q2_filepath.to_excel(writer, sheet_name='Question 2', index=False, header=False)

def question3(q3_filepath, excel_file_path):
    # Calling helper functions to solve question 3
    question3_helper_type(q3_filepath, excel_file_path)
    question3_helper_currency(q3_filepath, excel_file_path)
    question3_helper_country(q3_filepath, excel_file_path)

def question3_helper_country(q3_filepath, excel_file_path):
    italy = poland = bulgaria = uk = country_total = 0
    italy_balance = poland_balance = uk_balance = bulgaria_balance= country_total_balance = 0
    percent_sum = percent_sum_balance = 0
    column_j = q3_filepath.iloc[:, 9] #market
    column_m = q3_filepath.iloc[:, 12] # country
    column_k = q3_filepath.iloc[:, 10] # balance
    for (val_j, val_m, val_k) in zip(column_j, column_m, column_k):
        if not isinstance(val_j, int):
            continue
        if val_m == 'Italy':
            italy += val_j
            italy_balance += val_k
        elif val_m == 'UK':
            uk += val_j
            uk_balance += val_k
        elif val_m == 'Poland':
            poland += val_j
            poland_balance += val_k
        elif val_m == 'Bulgaria':
            bulgaria += val_j
            bulgaria_balance += val_k
        country_total += val_j
        country_total_balance += val_k
    q3_filepath.iat[19, 2] = italy
    percent_sum += italy / country_total
    q3_filepath.iat[19, 3] = italy / country_total

    q3_filepath.iat[20, 2] = uk
    percent_sum += uk / country_total
    q3_filepath.iat[20, 3] = uk / country_total

    q3_filepath.iat[21, 2] = poland
    percent_sum += poland / country_total
    q3_filepath.iat[21, 3] = poland / country_total

    q3_filepath.iat[22, 2] = bulgaria
    percent_sum += bulgaria / country_total
    q3_filepath.iat[22, 3] = bulgaria / country_total

    q3_filepath.iat[23, 2] = country_total
    q3_filepath.iat[23, 3] = percent_sum

    q3_filepath.iat[19, 4] = italy_balance
    percent_sum_balance += italy_balance / country_total_balance
    q3_filepath.iat[19, 5] = italy_balance / country_total_balance

    q3_filepath.iat[20, 4] = uk_balance
    percent_sum_balance += uk_balance / country_total_balance
    q3_filepath.iat[20, 5] = uk_balance / country_total_balance

    q3_filepath.iat[21, 4] = poland_balance
    percent_sum_balance += poland_balance / country_total_balance
    q3_filepath.iat[21, 5] = poland_balance / country_total_balance

    q3_filepath.iat[22, 4] = bulgaria_balance
    percent_sum_balance += bulgaria_balance / country_total_balance
    q3_filepath.iat[22, 5] = bulgaria_balance / country_total_balance

    q3_filepath.iat[23, 4] = country_total_balance
    q3_filepath.iat[23, 5] = percent_sum_balance

    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 3'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q3_filepath.to_excel(writer, sheet_name='Question 3', index=False, header=False)


def question3_helper_currency(q3_filepath, excel_file_path):
    EUR = GBP = BGN = PLN = currency_total = 0
    EUR_balance = GBP_balance = BGN_balance = PLN_balance = currency_total_balance = 0
    percent_sum = percent_sum_balance = 0
    column_j = q3_filepath.iloc[:, 9] #market
    column_l = q3_filepath.iloc[:, 11] # currency
    column_k = q3_filepath.iloc[:, 10] # balance
    for (val_j, val_l, val_k) in zip(column_j, column_l, column_k):
        if not isinstance(val_j, int):
            continue
        if val_l == 'EUR':
            EUR += val_j
            EUR_balance += val_k
        elif val_l == 'GBP':
            GBP += val_j
            GBP_balance += val_k
        elif val_l == 'BGN':
            BGN += val_j
            BGN_balance += val_k
        elif val_l == 'PLN':
            PLN += val_j
            PLN_balance += val_k
        currency_total += val_j
        currency_total_balance += val_k
    print(currency_total_balance)
    print(currency_total)

    q3_filepath.iat[12, 2] = EUR
    percent_sum += EUR / currency_total
    q3_filepath.iat[12, 3] = EUR / currency_total

    q3_filepath.iat[13, 2] = GBP
    percent_sum += GBP / currency_total
    q3_filepath.iat[13, 3] = GBP / currency_total

    q3_filepath.iat[14, 2] = BGN
    percent_sum += BGN / currency_total
    q3_filepath.iat[14, 3] = BGN / currency_total

    q3_filepath.iat[15, 2] = PLN
    percent_sum += PLN / currency_total
    q3_filepath.iat[15, 3] = PLN / currency_total

    q3_filepath.iat[16,2] = currency_total
    q3_filepath.iat[16,3] = percent_sum


    q3_filepath.iat[12, 4] = EUR_balance
    percent_sum_balance += EUR_balance / currency_total_balance
    q3_filepath.iat[12, 5] = EUR_balance / currency_total_balance

    q3_filepath.iat[13, 4] = GBP_balance
    percent_sum_balance += GBP_balance / currency_total_balance
    q3_filepath.iat[13, 5] = GBP_balance / currency_total_balance

    q3_filepath.iat[14, 4] = BGN_balance
    percent_sum_balance += BGN_balance / currency_total_balance
    q3_filepath.iat[14, 5] = BGN_balance / currency_total_balance

    q3_filepath.iat[15, 4] = PLN_balance
    percent_sum_balance += PLN_balance / currency_total_balance
    q3_filepath.iat[15, 5] = PLN_balance / currency_total_balance

    q3_filepath.iat[16,4] = currency_total_balance
    q3_filepath.iat[16,5] = percent_sum_balance

    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 3'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q3_filepath.to_excel(writer, sheet_name='Question 3', index=False, header=False)

def question3_helper_type(q3_filepath, excel_file_path):
    residential = commercial = land = industrial = other = type_total = 0
    residential_balance = commercial_balance = land_balance = industrial_balance = other_balance = type_total_balance = 0
    percent_sum = percent_sum_balance = 0
    column_j = q3_filepath.iloc[:, 9]
    column_i = q3_filepath.iloc[:, 8]
    column_k = q3_filepath.iloc[:, 10]
    for (val_j, val_i, val_k) in zip(column_j, column_i, column_k):
        if not isinstance(val_j, int) or val_j is None:
            continue
        if val_i == 'Residential':
            residential += val_j
            residential_balance += val_k
        elif val_i == 'Commercial':
            commercial += val_j
            commercial_balance += val_k
        elif val_i == 'Land':
            land += val_j
            land_balance += val_k
        elif val_i == 'Industrial':
            industrial += val_j
            industrial_balance += val_k
        else:
            other += val_j
            other_balance += val_k
        type_total += val_j
        type_total_balance += val_k
    print(type_total)
    print(type_total_balance)
    q3_filepath.iat[4,2] = residential
    percent_sum += residential / type_total
    q3_filepath.iat[4,3] = residential / type_total

    q3_filepath.iat[5,2] = commercial
    percent_sum += commercial / type_total
    q3_filepath.iat[5,3] = commercial / type_total

    q3_filepath.iat[6,2] = industrial
    percent_sum += industrial / type_total
    q3_filepath.iat[6,3] = industrial / type_total

    q3_filepath.iat[7,2] = land
    percent_sum += land / type_total
    q3_filepath.iat[7,3] = land / type_total

    q3_filepath.iat[8,2] = other
    percent_sum += other / type_total
    q3_filepath.iat[8,3] = other / type_total

    q3_filepath.iat[9,2] = type_total
    q3_filepath.iat[9,3] = percent_sum

    q3_filepath.iat[4,4] = residential_balance
    percent_sum_balance += residential_balance / type_total_balance
    q3_filepath.iat[4,5] = residential_balance / type_total_balance

    q3_filepath.iat[5,4] = commercial_balance
    percent_sum_balance += commercial_balance / type_total_balance
    q3_filepath.iat[5,5] = commercial_balance / type_total_balance

    q3_filepath.iat[6,4] = industrial_balance
    percent_sum_balance += industrial_balance / type_total_balance
    q3_filepath.iat[6,5] = industrial_balance / type_total_balance

    q3_filepath.iat[7,4] = land_balance
    percent_sum_balance += land_balance / type_total_balance
    q3_filepath.iat[7,5] = land_balance / type_total_balance

    q3_filepath.iat[8,4] = other_balance
    percent_sum_balance += other_balance / type_total_balance
    q3_filepath.iat[8,5] = other_balance / type_total_balance

    q3_filepath.iat[9,4] = type_total_balance
    q3_filepath.iat[9,5] = percent_sum_balance
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 3'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q3_filepath.to_excel(writer, sheet_name='Question 3', index=False, header=False)



def main():
    # Read Excel file
    excel_file_path = r'C:\Users\thoma\PycharmProjects\Balbec_Case_Study\Analyst Case Study Balbec Capital Python.xlsx'
    sheet1 = pd.read_excel(excel_file_path)
    sheet2 = pd.read_excel(excel_file_path, sheet_name=1)
    sheet3 = pd.read_excel(excel_file_path, sheet_name=2)


    question1(sheet1, excel_file_path)
    question2_real(sheet2, excel_file_path)
    question3(sheet3, excel_file_path)
if __name__ == "__main__":
    main()
