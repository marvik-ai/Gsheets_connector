# Google Drive and GSheets manager

This project allows you to interact with Google Drive and Google Sheets through a Python class called `GoogleDriveManager`. You can list files in a Google Drive folder, generate public links for files, and upload data (including images) to Google Sheets.

## TL;DR

### Setup
Load credentials from either a `.env` file or a `.json` file.

### Initialization
Create an instance of `GoogleDriveManager` using your credentials.

### Operations
- List files in a Google Drive folder.
- Retrieve the public link for a file.
- Upload a DataFrame to Google Sheets, with the ability to insert images.

## Installation

To get started, install the required dependencies:

```bash
pip install requirements.txt
```

## Usage

There are 3 possible ways to initialize:

1. If you use a .env file to load the credentials: <br>
```
manager = GoogleDriveManager(folder_id='your_drive_folder_id', use_dotenv=True)<br>
```
This will load the credentials from a .env file that contains the **GOOGLE_CREDENTIALS_JSON** variable.<br>

2. If you use a .json file with the credentials:<br>
```
manager = GoogleDriveManager(service_account_file='json_path.json', folder_id='your_drive_folder_id')<br>
```
This will load the credentials from a Google Service Account JSON file stored in your local file system.<br>

3. If you use environment variables directly on your system:<br>
First, set the environment variable in your terminal or system:<br>
```
export GOOGLE_CREDENTIALS_JSON='{your_json_content}'
```
Then, initialize the manager as follows:<br>
```
manager = GoogleDriveManager(folder_id='your_drive_folder_id')
```
This will load the credentials from the environment variable **GOOGLE_CREDENTIALS_JSON** that you exported on your system.<br>


### To list files from a drive folder
You can list the files from a specific Google Drive folder as follows:

```python
files_df = manager.list_files_in_folder()
```

To get the public link of a specific file in your Google Drive folder, use the following code:

### Get the public link of a file by its name

```python
file_name = "nombre_del_archivo_en_drive.ext"
drive_link = manager.get_drive_link(file_name)
```
### Upload a DataFrame to Google Sheets

If you want to upload a DataFrame to Google Sheets, you can do so with the following code. Make sure you have the Google Sheet ID and have shared edit permissions with the service account:

```python
spreadsheet_id = 'YOUR_SPREADSHEET_ID'
sheet_name = 'SHEET_NAME'
manager.create_sheet_with_data(spreadsheet_id=spreadsheet_id, sheet_name=sheet_name, df=df, image_col='image_col_name')
```

**Notes**

Make sure your Google Service Account has access to the Google Drive folder and Google Sheets you want to interact with.
For the image column in the DataFrame, the script will attempt to find the image in Google Drive and add it to the corresponding cell in Google Sheets.