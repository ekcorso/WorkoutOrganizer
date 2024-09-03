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


class SpreadsheetRow:
    original_name: str
    description: str
    skip: bool

    def __init__(self, original_name: str, description: str, skip: bool) -> None:
        self.original_name = original_name
        self.description = description
        self.skip = self.should_skip(skip)

    def __repr__(self) -> str:
        return f"SpreadsheetRow({self.original_name}, {self.description}, {self.skip})"

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
    spreadsheet: Spreadsheet, destination_folder_id: str, client: Client, translated_data: [SpreadsheetRow]
) -> None:
    """Copy all the sheets for spreadsheet and make each into a new spreadsheet"""
    sheets = spreadsheet.worksheets()
    for sheet in sheets:
            title = get_dest_spreadsheet_title(spreadsheet, sheet, translated_data)
        canary_cells = sheet.batch_get(["A1", "B2", "B4"])
        if is_valid_workout(canary_cells):
            dest_spreadsheet = create_new_spreadsheet(title, destination_folder_id, client)
            sheet.copy_to(dest_spreadsheet.id)
            dest_spreadsheet.del_worksheet(dest_spreadsheet.sheet1)


def should_process_spreadsheet(spreadsheet: Spreadsheet, translations: [SpreadsheetRow]) -> bool:
    """Check if translation sheet indicates that the workout should be skipped"""
    for translation in translations:
        if translation.original_name == spreadsheet.title:
            return not translation.skip


def is_valid_workout(canary_cells: list[list[list[str]]]) -> bool:
    """Check that the worksheet is not a blank workout template"""
    canary_cell_value = "" # A1

    if isinstance(canary_cells[0], list) and canary_cells[0]:
        canary_cell_value = canary_cells[0][0][0] if isinstance(canary_cells[0][0][0], str) else ""
    else:
        canary_cell_value = ""

    completely_blank = is_empty_3d_list(canary_cells)
    is_valid = False if (canary_cell_value == "Name: " or completely_blank) else True 
    return is_valid


def is_empty_3d_list(data: list[list[list[str]]]) -> bool:
   """Check if the 3D list is empty"""
   flattened_list = [item for sublist1 in data for sublist2 in sublist1 for item in sublist2]
   return not any(flattened_list)


def get_dest_spreadsheet_title(spreadsheet: Spreadsheet, worksheet: Worksheet, translated_data: [SpreadsheetRow]) -> str:
    """Create and return a title for the new spreadsheet"""
    tab_name = get_workout_description_for_worksheet(worksheet)
    source_title = spreadsheet.title
    translated_source_title = translate_workout_name(source_title, translated_data)
    new_spreadsheet_name = tab_name + " - " + translated_source_title
    return new_spreadsheet_name


def translate_workout_name(source_title: str, translated_data: [SpreadsheetRow]) -> str:
    """Translate the workout title from the source title to a supplied description"""
    translated_name = next((item.description for item in translated_data if item.original_name == source_title), None)
    return str(translated_name) if translated_name else str(source_title)


def get_client_name_list_from_spreadsheets(spreadsheets: [Spreadsheet]) -> [str]:
    """Return a list of titles for the given list of spreadsheets"""
    return [[sheet["name"]] for sheet in spreadsheets]


def create_workout_translation_spreadsheet(
    dest_folder_id: str, client: Client, current_client_files: [Spreadsheet]
) -> Spreadsheet:
    """Create a new spreadsheet in the destination folder with the title 'Workout Translation' and return the new spreadsheet
    with a list of all the client names in the origin folder."""
    spreadsheet = create_new_spreadsheet("Workout Translations", dest_folder_id, client)
    names = get_client_name_list_from_spreadsheets(current_client_files)
    sheet = spreadsheet.sheet1
    all_rows = ["Original Name", "Description"].append(names)
    sheet.append_rows(all_rows)
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
    translation_data = [SpreadsheetRow(row["Original Name"], row["Description"], row["Skip?"]) for row in translation_sheet.get_all_records()]

    needs_client_list = input("Do you want to create the client list? (y/n) ")
    if needs_client_list.lower() == "y":
        create_workout_translation_spreadsheet(
            destination_folder_id, client, spreadsheets_to_copy
        )


    for spreadsheet in track(spreadsheets_to_copy, "Copying..."):
        spreadsheet = client.open_by_key(spreadsheet["id"])
        if should_process_spreadsheet(spreadsheet, translation_data):
            separate_and_copy_all_sheets_to_folder(
                spreadsheet, destination_folder_id, client, translation_data
            )


if __name__ == "__main__":
    main()
