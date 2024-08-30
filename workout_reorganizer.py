"""Takes a set of workouts in a google sheets folder and iterates over them to extract each workout into a separate spreadsheet.

Requirements:
- Create an account in Google Cloud and enable Google Sheets API and Google Drive API
- Add a service account with the Editor role
- Share the source and destination folders with the service account email address and give it editor permission
- Create a Google Sheets API key and save it in the ~/.config/gspread/ directory
- Pip install the following: google-api-python-client, gspread
"""

import gspread

from gspread import Client
from gspread.spreadsheet import Spreadsheet
from gspread.worksheet import Worksheet

from rich.progress import track


class TranslationRow:
    original_name: str
    description: str
    skip: bool

    def __init__(self, original_name: str, description: str, skip: bool) -> None:
        self.original_name = original_name
        self.description = description
        self.skip = self.should_skip(skip)

    def __repr__(self) -> str:
        return f"TranslationRow({self.original_name}, {self.description}, {self.skip})"

    def should_skip(self, skip: str) -> bool:
        if skip:
            if skip.lower() == "y":
                return True
        return False


def fetch_list_of_files_in_folder(folder_id: str, client: Client) -> [Spreadsheet]:
    """Fetch a list of files in the given folder and return a list of Spreadsheet objects"""
    spreadsheets_in_folder = client.list_spreadsheet_files(
        title=None, folder_id=folder_id
    )
    valid_spreadsheets = list(filter(lambda n: "Workout Translations" not in n["name"], spreadsheets_in_folder))
    return valid_spreadsheets


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
            title = get_dest_spreadsheet_title(spreadsheet, sheet, client)
            dest_spreadsheet = create_new_spreadsheet(title, destination_folder_id, client)
            sheet.copy_to(dest_spreadsheet.id)
            remove_sheet1_from_spreadsheet(dest_spreadsheet)


def should_process_spreadsheeet(spreadsheet: Spreadsheet, client: Client) -> bool:
    """Check if translation sheet indicates that the workout should be skipped"""
    # TODO: Replace the hardcoded test name with the actual name of the translation spreadsheet
    translation_sheet = client.open("Workout Translations - TEST").sheet1
    translation_row = translation_sheet.find(spreadsheet.title).row
    skip_cell = translation_sheet.cell(translation_row, 3).value
    if skip_cell:
        if skip_cell.lower() == "y":
            return False
    return True


def is_valid_workout(worksheet: Worksheet) -> bool:
    """Check that the worksheet is not a blank workout template"""
    canary_cell_value = worksheet.acell("A1").value
    completely_blank = not any(worksheet.get_all_values())
    is_valid = False if (canary_cell_value == "Name: " or completely_blank) else True 
    return is_valid


def get_dest_spreadsheet_title(spreadsheet: Spreadsheet, worksheet: Worksheet, client: Client) -> str:
    """Create and return a title for the new spreadsheet"""
    tab_name = get_workout_description_for_worksheet(worksheet)
    source_title = spreadsheet.title
    translated_source_title = translate_workout_name(source_title, client)
    new_spreadsheet_name = tab_name + " - " +translated_source_title
    return new_spreadsheet_name


def translate_workout_name(source_title: str, client: Client) -> str:
    """Translate the workout title from the source title to a supplied description"""
    # TODO: Replace the hardcoded test name with the actual name of the translation spreadsheet
    translation_sheet = client.open("Workout Translations - TEST").sheet1
    translation_row = translation_sheet.find(source_title).row
    translated_name = translation_sheet.cell(translation_row, 2).value
    return str(translated_name) if translated_name else str(source_title)


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


def get_workout_description_for_worksheet(worksheet: Worksheet) -> str:
    """Return the description of the workout from the spreadsheet"""
    newer_description = worksheet.acell("B2").value
    older_description = worksheet.acell("B4").value
    canary_cell = worksheet.acell("A1").value # This cell will be blank if the workout is an older format
    ret = older_description if not canary_cell else newer_description
    return ret


def main() -> None:
    print("Hint: folder IDs are in the URL of the folder in Google Drive.")
    source_folder_id = str(input("Please enter the source folder ID: "))
    destination_folder_id = str(input("Please enter the destination folder ID: "))
    client = gspread.service_account()

    spreadsheets_to_copy = fetch_list_of_files_in_folder(source_folder_id, client)
    
    translation_sheet = client.open("Workout Translations - TEST").sheet1
    translation_data = [TranslationRow(row["Original Name"], row["Description"], row["Skip?"]) for row in translation_sheet.get_all_records()]

    needs_client_list = input("Do you want to create the client list? (y/n) ")
    if needs_client_list.lower() == "y":
        create_workout_translation_spreadsheet(
            source_folder_id, destination_folder_id, client
        )


    for spreadsheet in track(spreadsheets_to_copy, "Copying..."):
        spreadsheet = client.open_by_key(spreadsheet["id"])
        if should_process_spreadsheeet(spreadsheet, client):
            separate_and_copy_all_sheets_to_folder(
                spreadsheet, destination_folder_id, client
            )


if __name__ == "__main__":
    main()
