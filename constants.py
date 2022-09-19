FILE_NAME = "One Form - Phone Backup (3)(1-2).xlsx"
SHEET_NAME = "Sheet1"


# Values are Column names as they are on Microsoft Form

unknown = "Unknown"
i_number = "I-Number"
department = "What area is this issue about?"
notes = "Contact Notes"

COLUMNS = [
    unknown,
    i_number,
    department, 
    notes
    ] 

email_dept_match = {
    "Pathway Student Unknown" : ["unknownpathway@unknown.com","PathwayOnlineStudent"],
    "On Campus Student Unknown" : ["campusunknown@unknown.com", "OnTrackStudent"], 
    "Online Student Unknown" : ["onlineunknown@unknown.com", "OnlineStudent"],
    "Online Prospective Student Unknown" : ["onlineprospectiveunknown@unknown.com", "OnlineProspectiveStudent"],
    "Community Member Unknown" : ["communitymember@unknown.com", "Community Member"],
    "Parent/Other Relative Unknown" : ["Parentunknown@unknown.com", "Parent/Other Relative"],
    "Alumni/Former BYUI Student Unknown" : ["alumniunknown@unknown.com", "Alumni"],
    "Employee/Faculty Unknown" : ["employeeunknown@unknown.com", "Student Services"],
    "Vendor ." : ["vendor@unknown.com", "Vendor"],
    "Church Leader Unknown" : ["churchleader@unknown.com", "Bishop/Stake Pres"],
    "Missionary Unknown" : ["missionaryunknown@unknown.com", "MissionaryDeferredStudent"],
}

contact_type = {
    "Pathway Student Unknown" : "PathwayOnlineStudent",
    "On Campus Student Unknown" : "OnTrackStudent",
    "Online Student Unknown" : "OnlineStudent",
    "Online Prospective Student Unknown" : "OnlineProspectiveStudent",
    "Community Member Unknown" : "Community Member",
    "Parent/Other Relative Unknown" : "Parent/Other Relative",
    "Alumni/Former BYUI Student Unknown" : "Alumni",
    "Employee/Faculty Unknown" : "Student Services",
    "Vendor ." : "Vendor",
    "Church Leader Unknown" : "Bishop/Stake Pres",
    "Missionary Unknown" : "MissionaryDeferredStudent"
}


# Values are One Form attributes in TD

ATTRIBUTES ={
    unknown : "Requestor",
    i_number : "I-Number (Back Up Phone Form)",
    department : "Office List 1",
    notes : "Contact Notes",
}

TD_ACTION = "Action Options - Phones"
ACTIONS = ["Which action did you use to help the contact?"] 
ACTIONS += [f"Which action did you use to help the contact?{n}" for n in range (2, 11)]
ACTION_DICT = {column : TD_ACTION for column in ACTIONS}




