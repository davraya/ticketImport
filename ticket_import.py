from csv import excel
from types import NoneType
import constants
import pandas as pd 




# form is a pandas.DataFrame
form = pd.read_excel(
                    open(constants.FILE_NAME, "rb"), 
                    sheet_name=constants.SHEET_NAME
                    )


# col = form.get(constants.I_NUMBER)
# print(col.array[2])

excel_sheet = {}




for column_name in constants.COLUMNS:
    # column_data is a pandas.Series
    column_data = form.get(column_name)
    td_attribute = constants.ATTRIBUTES[column_name]

    excel_column_data = list()
    for cell in column_data:
        if column_name == constants.unknown and str(cell) == "nan":
             excel_column_data.append(f"Community Member Unknown")
            # if str(cell) == "nan":
            #     excel_column_data.append(f"Community Member Unknown <communitymember@unknown.com>")
            # else:
            #     excel_column_data.append(f"{cell} <{constants.email_match[cell]}>")
        else:
            excel_column_data.append(cell)
    excel_sheet[td_attribute] = excel_column_data

number_of_rows = form.index.stop

requestor_column = excel_sheet["Requestor"]
excel_sheet["Acct/Dept"] = [constants.email_dept_match[requestor][1] for requestor in requestor_column]
excel_sheet["Title"] = ["One Form - Phone"] * number_of_rows 
excel_sheet["Ticket Type"] = ["(BSC) / phone call"] * number_of_rows

for i, requestor in enumerate(requestor_column):
    email = constants.email_dept_match[requestor][0]
    excel_sheet["Requestor"][i] = f"{requestor} <{email}>" 










actions = constants.ACTION_DICT.keys()

action_choices = [""] * number_of_rows


for action in actions:
    column_data = form.get(action)

    if type(column_data) != NoneType:
        for i, cell in enumerate(column_data):
            if str(cell) != "nan":
                action_choices[i] = cell

excel_sheet[constants.TD_ACTION] = action_choices

data_frame = pd.DataFrame(excel_sheet)
data_frame.to_excel("TicketImport.xlsx")

print("The program completed")


            
    # td_attribute = constants.ACTION_DICT[action]

    # if "Nan" in column_data:
    #     print(action, "not selected")





# print(excel_sheet)