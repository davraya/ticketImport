from types import NoneType
import constants
import pandas as pd 
import numpy as np



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
        excel_column_data.append(cell)
    
    excel_sheet[td_attribute] = excel_column_data
         
actions = constants.ACTION_DICT.keys()

number_of_rows = form.index.stop

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