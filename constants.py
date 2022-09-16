FILE_NAME = "One Form - Phone Backup (2)(1-4).xlsx"
SHEET_NAME = "Sheet1"


# Values are Column names as they are on Microsoft Form

contact = "Contact's Name"
i_number = "I-Number"
department = "What area is this issue about?"
notes = "Contact Notes"

ACTIONS = [f"Which action did you use to help the contact?{n}" for n in range (2, 11)]



COLUMNS = [
    contact,
    i_number,
    department, 
    notes
    ] 

# Values are One Form attributes names

ATTRIBUTES ={
    contact : "Requestor",
    i_number : "I-Number (Back Up Phone Form)",
    department : "Office List 1",
    notes : "Contact Notes"
}

TD_ACTION = "Action Options - Phones"

ACTION_DICT = {column : TD_ACTION for column in ACTIONS}


