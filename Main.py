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
            print(index)
            q1_filepath.iat[int(index), 2] = temp
            print("post update")

    # Use ExcelWriter to append to the original Excel file
    with pd.ExcelWriter(excel_file_path, engine='openpyxl', mode='a') as writer:
        # Remove the existing sheet if it exists
        try:
            writer.book.remove(writer.sheets['Question 1'])
        except KeyError:
            pass

        # Write the DataFrame to the existing or new sheet
        q1_filepath.to_excel(writer, sheet_name='Question 1', index=False, header=False)



def main():
    # Read Excel file
    excel_file_path = r'C:\Users\thoma\PycharmProjects\Balbec_Case_Study\Analyst Case Study Balbec Capital Python.xlsx'
    sheet1 = pd.read_excel(excel_file_path)
    sheet2 = pd.read_excel(excel_file_path, sheet_name=1)
    sheet3 = pd.read_excel(excel_file_path, sheet_name=2)

    question1(sheet1, excel_file_path)

if __name__ == "__main__":
    main()
