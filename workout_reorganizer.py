"""Takes a set of workouts in a google sheets folder and iterates over them to extract each workout into a separate spreadsheet.

Requirements:
- Create an account in Google Cloud and enable Google Sheets API and Google Drive API
- Add a service account with the Editor role
- Share the source and destination folders with the service account email address and give it editor permission
- Create a Google Sheets API key and save it in the ~/.config/gspread/ directory
- Pip install the following: google-api-python-client, gspread
"""

import gspread  # type: ignore[import-untyped]

from concurrent.futures import ThreadPoolExecutor, as_completed
from gspread import Client
from gspread import Spreadsheet
from gspread import Worksheet

from rich.progress import track

from tenacity import retry, stop_after_attempt, wait_exponential

common_retry = retry(
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=1, max=60),
)

class SpreadsheetRow:
    original_name: str
    description: str
    skip: bool

    def __init__(self, original_name: str, description: str, skip: str) -> None:
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


@common_retry
def fetch_list_of_files_in_folder(folder_id: str, client: Client) -> list[Spreadsheet]:
    """Fetch a list of files in the given folder and return a list of Spreadsheet objects"""
    spreadsheets_in_folder = client.list_spreadsheet_files(
        title=None, folder_id=folder_id
    )
    valid_spreadsheets = list(
        filter(
            lambda n: "Workout Translations" not in n["name"], spreadsheets_in_folder
        )
    )
    return valid_spreadsheets


@common_retry
def create_new_spreadsheet(
    title: str, dest_folder_id: str, client: Client
) -> Spreadsheet:
    """Create a new spreadsheet in the destination folder with the given title and return the new spreadsheet"""
    sheet = client.create(title, folder_id=dest_folder_id)
    return sheet


@common_retry
def share_spreadsheet(sheet: Spreadsheet) -> None:
    """Share the spreadsheet with the service account email address"""
    sheet.share("ekcorso@gmail.com", perm_type="user", role="writer", notify=False)


def create_and_share_new_spreadsheet(title: str, dest_folder_id: str, client: Client) -> Spreadsheet:
    """Create a new spreadsheet, share it, and return it"""
    sheet = create_new_spreadsheet(title, dest_folder_id, client)
    share_spreadsheet(sheet)
    return sheet


def separate_and_copy_all_sheets_to_folder(
    spreadsheet: Spreadsheet,
    destination_folder_id: str,
    client: Client,
    translated_data: list[SpreadsheetRow],
) -> None:
    """Copy all the sheets for spreadsheet and make each into a new spreadsheet in parallel"""
    sheets = spreadsheet.worksheets()

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []

        for sheet in sheets:
            canary_cells = get_canary_cells(sheet)
            previous_description = get_workout_description(canary_cells)
            if is_valid_workout(canary_cells, previous_description):
                futures.append(
                    executor.submit(
                        process_sheet,
                        sheet,
                        spreadsheet,
                        destination_folder_id,
                        client,
                        previous_description,
                        translated_data,
                    )
                )

        for future in as_completed(futures):
            future.result()


@common_retry
def get_canary_cells(sheet: Worksheet) -> list[list[list[str]]]:
    """Get the values of the canary cells from the sheet"""
    return sheet.batch_get(["A1", "B1", "B2", "B3", "B4"])


def process_sheet(
    sheet: Worksheet,
    spreadsheet: Spreadsheet,
    destination_folder_id: str,
    client: Client,
    previous_description: str,
    translated_data: list[SpreadsheetRow],
) -> None:
    """Process a single sheet and copy it to a new spreadsheet in the destination folder"""
    title = get_dest_spreadsheet_title(
        spreadsheet, previous_description, translated_data
    )
    dest_spreadsheet = create_and_share_new_spreadsheet(title, destination_folder_id, client)
    new_worksheet_data = copy_workout_to_new_spreadsheet(dest_spreadsheet.id, sheet)
    new_worksheet = get_worksheet_from_data(new_worksheet_data, dest_spreadsheet)
    rename_worksheet(new_worksheet, title)
    delete_client_data(get_canary_cells(sheet), new_worksheet)
    delete_empty_sheet1(dest_spreadsheet)


@common_retry
def copy_workout_to_new_spreadsheet(dest_spreadsheet_id: str, sheet: Worksheet) -> dict[str, str]:
    """Copy the worksheet to the new spreadsheet and return the new worksheet json data"""
    new_worksheet_data = sheet.copy_to(dest_spreadsheet_id)
    return new_worksheet_data


@common_retry
def get_worksheet_from_data(data: dict[str, str], dest_spreadsheet: Spreadsheet) -> Worksheet:
    """Get the Worksheet object from the json data, then return it"""
    return dest_spreadsheet.get_worksheet_by_id(data["sheetId"])


@common_retry
def rename_worksheet(worksheet: Worksheet, new_title: str) -> None:
    """Rename the worksheet to the new title"""
    worksheet.update_title(new_title)


@common_retry
def delete_empty_sheet1(spreadsheet: Spreadsheet) -> None:
    """Delete the empty first sheet in the spreadsheet that is created by default"""
    spreadsheet.del_worksheet(spreadsheet.sheet1)


def should_process_spreadsheet(
    spreadsheet: Spreadsheet, translations: list[SpreadsheetRow]
) -> bool:
    """Check if translation sheet indicates that the workout should be skipped"""
    for translation in translations:
        if translation.original_name == spreadsheet.title:
            return not translation.skip
    return False  # If no translation is found, skip the workout


def is_valid_workout(
    canary_cells: list[list[list[str]]], previous_description: str
) -> bool:
    """Check that the worksheet is not a blank workout template"""
    canary_cell_value = ""  # A1

    if isinstance(canary_cells[0], list) and canary_cells[0]:
        canary_cell_value = (
            canary_cells[0][0][0] if isinstance(canary_cells[0][0][0], str) else ""
        )
    else:
        canary_cell_value = ""

    completely_blank = not any(flatten_3d_list(canary_cells))
    is_foundation_workout = "Foundation 1" in previous_description
    is_valid = (
        False
        if (canary_cell_value == "Name: " or is_foundation_workout or completely_blank)
        else True
    )
    return is_valid


@common_retry
def delete_client_data(canary_cells: list[list[list[str]]], sheet: Worksheet):
    """Delete client name from the worksheet"""
    new_name_location = 1  # B1
    old_name_location = 3  # B3

    old_name = get_value_at_location(canary_cells, old_name_location)
    new_name = get_value_at_location(canary_cells, new_name_location)

    if new_name:
        # Delete the client name and any additional private data in the 1st row
        sheet.batch_update([{"range": "B1:D1", "values": [["", "", ""]]}])
    elif old_name:
        sheet.update_cell(3, 2, "")


def flatten_3d_list(data: list[list[list[str]]]) -> list[str]:
    """Flatten a 3D list"""
    return [item for sublist1 in data for sublist2 in sublist1 for item in sublist2]


def get_dest_spreadsheet_title(
    spreadsheet: Spreadsheet,
    previous_description: str,
    translated_data: list[SpreadsheetRow],
) -> str:
    """Create and return a title for the new spreadsheet"""
    source_title = spreadsheet.title
    translated_source_title = translate_workout_name(source_title, translated_data)
    new_spreadsheet_name = previous_description + " - " + translated_source_title
    return new_spreadsheet_name.title()


def translate_workout_name(
    source_title: str, translated_data: list[SpreadsheetRow]
) -> str:
    """Translate the workout title from the source title to a supplied description"""
    translated_name = next(
        (
            item.description
            for item in translated_data
            if item.original_name == source_title
        ),
        None,
    )
    return str(translated_name) if translated_name else str(source_title)


def get_client_name_list_from_spreadsheets(
    spreadsheets: list[Spreadsheet],
) -> list[list[str]]:
    """Return a list of titles for the given list of spreadsheets"""
    return [[sheet["name"]] for sheet in spreadsheets]


def create_workout_translation_spreadsheet(
    dest_folder_id: str, client: Client, current_client_files: list[Spreadsheet]
) -> Spreadsheet:
    """Create a new spreadsheet in the destination folder with the title 'Workout Translations' and return the new spreadsheet
    with a list of all the client names in the origin folder."""
    spreadsheet = create_and_share_new_spreadsheet(
        "Workout Translations", dest_folder_id, client
    )  # Create in dest folder so that the client's original folder is never modified

    sheet = append_rows_to_sheet(spreadsheet, current_client_files)
    return sheet


@common_retry
def append_rows_to_sheet(spreadsheet: Spreadsheet, current_client_files: list[Spreadsheet]) -> Worksheet:
    """Append rows to the given worksheet"""
    sheet = spreadsheet.sheet1

    names = get_client_name_list_from_spreadsheets(current_client_files)
    all_rows = [["Original Name", "Description", "Skip?"]] + names

    sheet.append_rows(all_rows)

    return sheet


def get_workout_description(canary_cells: list[list[list[str]]]) -> str:
    """Return the description of the workout from the spreadsheet"""
    canary_cell_location = (
        0  # A1: this cell will be blank if the workout is an older format
    )
    new_description_location = 2  # B2
    old_description_location = 4  # B4

    canary_cell_value = get_value_at_location(canary_cells, canary_cell_location)
    newer_description = get_value_at_location(canary_cells, new_description_location)
    older_description = get_value_at_location(canary_cells, old_description_location)

    return older_description if not canary_cell_value else newer_description


def get_value_at_location(canary_cells: list[list[list[str]]], location: int) -> str:
    """Return the value at the given location in the 3d list"""
    if canary_cells[location] and isinstance(canary_cells[location], list):
        return (
            canary_cells[location][0][0]
            if isinstance(canary_cells[location][0][0], str)
            else ""
        )
    else:
        return ""


def main() -> None:
    print("Hint: folder IDs are in the URL of the folder in Google Drive.")
    source_folder_id = str(input("Please enter the source folder ID: "))
    destination_folder_id = str(input("Please enter the destination folder ID: "))
    client = gspread.service_account()

    spreadsheets_to_copy = fetch_list_of_files_in_folder(source_folder_id, client)
    translation_sheet = client.open("Workout Translations - TEST").sheet1
    translation_data = [
        SpreadsheetRow(row["Original Name"], row["Description"], row["Skip?"])
        for row in translation_sheet.get_all_records()
    ]

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
