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


def iterate_over_all_spreadsheets_in_folder(folder_id: str) -> None:
    """Iterate over all the spreadsheets in the given folder and call the function to copy each sheet to a new spreadsheet"""
    folder = [str]  # replace this with the actual Google Drive folder object
    for spreadsheet in folder:
        separate_and_copy_all_sheets_to_folder(spreadsheet)


def separate_and_copy_all_sheets_to_folder(spreadsheet: gspread.spreadsheet.Spreadsheet) -> None:
    """Copy all the sheets for spreadsheet and make each it's own spreasheet"""
    client = gspread.service_account()
    destination_folder_id = ""  # replace this with the ID of the folder where the new spreadsheets should be saved
    sheets = spreadsheet.worksheets()
    for sheet in sheets:
        title = get_dest_spreadsheet_title(spreadsheet, sheet)
        dest_spreadsheet_id = create_new_spreadsheet_with_title(title)
        sheet.copy_to(dest_spreadsheet_id)


def get_dest_spreadsheet_title(spreadsheet: gspread.spreadsheet.Spreadsheet, worksheet: gspread.worksheet.Worksheet) -> str:
    """Create and return a title for the new spreadsheet based on 1) the name of the original spreadsheet and 2) the name of the sheet within that spreadsheet that is being copied"""
    client = gspread.service_account()
    new_spreadsheet_name = spreadsheet.title + " - " + worksheet.title
    return new_spreadsheet_name


def main() -> None:
    pass


if __name__ == "__main__":
    main()
