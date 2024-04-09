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

from rich.progress import track


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
    sheet.share("ekcorso@gmail.com", perm_type="user", role="writer", notify=False)
    return sheet


def separate_and_copy_all_sheets_to_folder(
    spreadsheet: Spreadsheet, destination_folder_id: str, client: Client
) -> None:
    """Copy all the sheets for spreadsheet and make each into a new spreadsheet"""
    sheets = spreadsheet.worksheets()
    for sheet in sheets:
        if is_valid_workout(sheet):
            title = get_dest_spreadsheet_title(spreadsheet, sheet)
            dest_spreadsheet = create_new_spreadsheet(title, destination_folder_id, client)
            sheet.copy_to(dest_spreadsheet.id)
            remove_sheet1_from_spreadsheet(dest_spreadsheet)


def is_valid_workout(worksheet: Worksheet) -> bool:
    """Check that the worksheet is not a blank workout template"""
    canary_cell_value = worksheet.acell("A1").value
    completely_blank = not any(worksheet.get_all_values())
    is_valid = False if (canary_cell_value == "Name: " or completely_blank) else True 
    return is_valid


def get_dest_spreadsheet_title(spreadsheet: Spreadsheet, worksheet: Worksheet) -> str:
    """Create and return a title for the new spreadsheet"""
    new_spreadsheet_name = spreadsheet.title + " - " + worksheet.title
    return new_spreadsheet_name


def remove_sheet1_from_spreadsheet(spreadsheet: Spreadsheet) -> None:
    """Remove the default 'Sheet1' from the spreadsheet"""
    sheet1 = spreadsheet.worksheet("Sheet1")
    spreadsheet.del_worksheet(sheet1)


def get_client_name_list_from_spreadsheets(spreadsheets: [Spreadsheet]) -> [str]:
    """Return a list of titles for the given list of spreadsheets"""
    return [[sheet["name"]] for sheet in spreadsheets]


def create_workout_translation_spreadsheet(
    origin_folder_id: str, dest_folder_id: str, client: Client
) -> Spreadsheet:
    """Create a new spreadsheet in the destination folder with the title 'Workout Translation' and return the new spreadsheet
    with a list of all the client names in the origin folder."""
    spreadsheet = create_new_spreadsheet("Workout Translations", dest_folder_id, client)
    current_client_files = fetch_list_of_files_in_folder(origin_folder_id, client)
    names = get_client_name_list_from_spreadsheets(current_client_files)
    sheet = spreadsheet.get_worksheet(0)
    sheet.append_row(["Original Name", "Description"])
    sheet.append_rows(names)
    return sheet


def main() -> None:
    print("Hint: folder IDs are in the URL of the folder in Google Drive.")
    source_folder_id = str(input("Please enter the source folder ID: "))
    destination_folder_id = str(input("Please enter the destination folder ID: "))
    client = gspread.service_account()

    spreadsheets_to_copy = fetch_list_of_files_in_folder(source_folder_id, client)

    needs_client_list = input("Do you want to create the client list? (y/n) ")
    if needs_client_list.lower() == "y":
        create_workout_translation_spreadsheet(
            source_folder_id, destination_folder_id, client
        )


    for spreadsheet in track(spreadsheets_to_copy, "Copying..."):
        spreadsheet = client.open_by_key(spreadsheet["id"])
        separate_and_copy_all_sheets_to_folder(
            spreadsheet, destination_folder_id, client
        )


if __name__ == "__main__":
    main()
