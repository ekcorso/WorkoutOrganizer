import gspread  # type: ignore[import-untyped]

from concurrent.futures import ThreadPoolExecutor, as_completed
from gspread import Client
from gspread import Spreadsheet
from gspread import Worksheet

from rich.progress import track

from tenacity import retry, stop_after_attempt, wait_exponential

from spreadsheet import SpreadsheetRow

def main() -> None:
    print("Hint: folder IDs are in the URL of the folder in Google Drive.")
    source_folder_id = str(input("Please enter the source folder ID: "))
    destination_folder_id = str(input("Please enter the destination folder ID: "))
    translation_sheet_id = str(input("Please enter the Translation sheet ID: "))

    try:
        client = get_client()
        spreadsheets_to_copy = fetch_list_of_files_in_folder(source_folder_id, client)
    except Exception as e:
        print(f"An error occured during setup: {e}")
        exit()

    try:
        translation_sheet = open_spreadsheet_by_key(translation_sheet_id, client).sheet1
        translation_data = get_translation_data(translation_sheet)
    except Exception as e:
        print(f"An error occured while setting up the translation sheet: {e}")
        exit()

    needs_client_list = input("Do you want to create the client list? (y/n) ")
    if needs_client_list.lower() == "y":
        create_workout_translation_spreadsheet(
            destination_folder_id, client, spreadsheets_to_copy
        )
        exit()

    for spreadsheet in track(spreadsheets_to_copy, "Copying..."):
        try:
            open_spreadsheet = open_spreadsheet_by_key(spreadsheet["id"], client)
        except Exception as e:
            spreadsheet_title = spreadsheet["name"]
            print(
                f"An error occured while opening the spreadsheet {spreadsheet_title} : {e}"
            )
            continue
        if should_process_spreadsheet(open_spreadsheet, translation_data):
            separate_and_copy_all_sheets_to_folder(
                open_spreadsheet, destination_folder_id, client, translation_data
            )


if __name__ == "__main__":
    main()
