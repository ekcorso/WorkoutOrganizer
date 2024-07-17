# Workout Organizer

**WorkoutOrganizer** processes workout spreadsheets to strip them of identifying client data and rename them according to their purpose so they can be used as workout templates.

## Introduction

This script takes a set of workout spreadsheets in a Google Sheets folder, iterates over them, extracts each workout into a new spreadsheet, and names the spreadsheet according to the workout's content.

The workouts are renamed with a translation spreadsheet that allows the coach to translate between a client's name and their workouts' content. The user is guided through the creation of this worksheet during the execution of the script.

## Requirements:
- Create an account in Google Cloud and enable Google Sheets API and Google Drive API
- Add a service account with the Editor role
- Share the source and destination folders with the service account email address and give that account editor permission
- Create a Google Sheets API key and save it in your ~/.config/gspread/ directory
- Install libraries:
 '''bash
pip install google-api-python-client gspread
'''

