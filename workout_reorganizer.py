## Workout Reorganizer Script
## Takes a set of workouts in a google sheets folder and iterate over them to extract each workout into a separate spreadsheet.

# Steps:
# Enable the Google Sheets API
# Fetch a list of files in the google sheets directory
# Iterate over all files in the google sheets directory
# For each file, iterate over each sheet
# Copy each sheet into a new spreadhseet by itself and save it in the appropriate directory

"""
Requirements:
- Create an account in Google Cloud and enable Google Sheets API and Google Drive API
- Add a sevice account with the Editor role
- Create a Google Sheets API key and save it in the ~/.config/gspread/ directory
- Pip install the following: google-api-python-client, gspread
# Confirm whether these are needed: google-auth-oauthlib, google-auth-oauthlib
"""


import os
import requests
import gspread


# First pass, try referencing a sheet with a known ID. Once I have enabled the Google Drive API as well I'll be able to fetch a list of the spreadsheet IDs that are in a given folder and work with each of them


def create_new_spreadsheet_with_title(title: str) -> str:
    """Create a new spreadsheet with the given title and return the ID of the new spreadsheet"""
    client = gspread.service_account()
    sheet = client.create(title)
    sheet.share("ekcorso@gmail.com", perm_type="user", role="writer")
    return sheet.id




def seperate_and_copy_all_sheets_to_folder(spreadsheet_id: str) -> None:
    """Copy all the sheets for spreadsheet and make each it's own spreashee"""t
    client = gspread.service_account()
    destination_folder_id = ""  # replace this with the ID of the folder where the new spreadsheets should be saved
    spreadsheet = client.open_by_key(spreadsheet_id)
    sheets = spreadsheet.worksheets()
    for sheet in sheets:
        title = get_dest_spreadsheet_title(spreadsheet_id)
        dest_spreadsheet_id = create_new_spreadsheet_with_title(title)
        sheet.id.copy_to(dest_spreadsheet_id)


def get_dest_spreadsheet_title(spreadsheet_id: str) -> str:
    """Create and return a title for the new spreadsheet based on 1) the name of the original spreadsheet and 2) the name of the sheet within that spreadsheet that is being copied"""
    client = gspread.service_account()
    # replace this with the spreadsheet_id passed in above
    sample_sheet_id = "1XmCCekBLcKKBMsc7pZ2egTmW7--6NmR9rf8jn96YzwY"
    spreadsheet = client.open_by_key(sample_sheet_id)
    sheets = spreadsheet.worksheets()
    new_spreadsheet_name = spreadsheet.title + " - " + sheets[0].title
    return new_spreadsheet_name


def main() -> None:
    pass


if __name__ == "__main__":
    main()
