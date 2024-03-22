"""Takes a set of workouts in a google sheets folder and iterates over them to extract each workout into a separate spreadsheet.

Requirements:
- Create an account in Google Cloud and enable Google Sheets API and Google Drive API
- Add a sevice account with the Editor role
- Share the source and destination folders with the service account email address and give it editor permission
- Create a Google Sheets API key and save it in the ~/.config/gspread/ directory
- Pip install the following: google-api-python-client, gspread
"""

import gspread

from gspread import Client
from gspread.spreadsheet import Spreadsheet
from gspread.worksheet import Worksheet



def fetch_list_of_files_in_folder(folder_id: str, client: Client) -> [Spreadsheet]:
    """Fetch a list of files in the given folder and return a list of Spreadsheet objects"""
    spreadsheets_in_folder = client.list_spreadsheet_files(
        title=None, folder_id=folder_id
    )
    return spreadsheets_in_folder


def create_new_spreadsheet(
    title: str, dest_folder_id: str, client: Client
) -> Spreadsheet:
    """Create a new spreadsheet in the destination folder with the given title and return the new spreadsheet"""
    sheet = client.create(title, folder_id=dest_folder_id)
    sheet.share("ekcorso@gmail.com", perm_type="user", role="writer")
    return sheet


def separate_and_copy_all_sheets_to_folder(
    spreadsheet: Spreadsheet, destination_folder_id: str, client: Client
) -> None:
    """Copy all the sheets for spreadsheet and make each into a new spreadsheet"""
    sheets = spreadsheet.worksheets()
    for sheet in sheets:
        title = get_dest_spreadsheet_title(spreadsheet, sheet)
        dest_spreadsheet = create_new_spreadsheet(title, destination_folder_id, client)
        sheet.copy_to(dest_spreadsheet.id)


def get_dest_spreadsheet_title(spreadsheet: Spreadsheet, worksheet: Worksheet) -> str:
    """Create and return a title for the new spreadsheet"""
    new_spreadsheet_name = spreadsheet.title + " - " + worksheet.title
    return new_spreadsheet_name


def main() -> None:
    print("Hint: folder IDs are in the URL of the folder in Google Drive.")
    source_folder_id = str(input("Please enter the source folder ID: "))
    destination_folder_id = str(input("Please enter the destination folder ID: "))
    client = gspread.service_account()

    spreadsheets_to_copy = fetch_list_of_files_in_folder(source_folder_id, client)

    print("Copying sheets now...")
    for spreadsheet in spreadsheets_to_copy:
        spreadsheet = client.open_by_key(spreadsheet["id"])
        separate_and_copy_all_sheets_to_folder(
            spreadsheet, destination_folder_id, client
        )


if __name__ == "__main__":
    main()
