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
- Create an account in Google Cloud and enable Google Sheets API
- Add a sevice account with the Editor role
- Create a Google Sheets API key and save it in the ~/.config/gspread/ directory
- Pip install the following: google-api-python-client, gspread
# Confirm whether these are needed: google-auth-oauthlib, google-auth-oauthlib
"""


import os
import requests
import gspread


# First pass, try referencing a sheet with a known ID. Once I have enabled the Google Drive API as well I'll be able to fetch a list of the spreadsheet IDs that are in a given folder and work with each of them


client = gspread.service_account()
sample_sheet_id = "1XmCCekBLcKKBMsc7pZ2egTmW7--6NmR9rf8jn96YzwY"

spreadsheet = client.open_by_key(sample_sheet_id)
sheets = spreadsheet.worksheets()  # Now we can ID sheets by their index
new_spreadsheet_name = spreadsheet.title + " - " + sheets[0].title
# client.create(new_spreadsheet_nam) # This will create a new spreadsheet with the name provided, but requires google drive API to be enabled


def main() -> None:
    pass


if __name__ == "__main__":
    main()
