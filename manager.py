import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import json
from dotenv import load_dotenv


class GoogleDriveManager:
    """
    A class to manage Google Drive and Google Sheets operations, including listing files,
    generating file links, and uploading data to Google Sheets.

    Attributes:
        creds (google.oauth2.credentials.Credentials): Google API credentials.
        drive_service (googleapiclient.discovery.Resource): Google Drive service object.
        gc (gspread.client.Client): Google Sheets service object.
        folder_id (str): The Google Drive folder ID used for file operations.
    """
    
    def __init__(self, service_account_file: str, folder_id: str):
        """
        Initializes the GoogleDriveManager with Google API credentials and Drive folder ID.

        Args:
            service_account_file (str): Path to the Google Service Account JSON file.
            folder_id (str): Google Drive folder ID for file operations.
        """
        SCOPES = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        self.creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
        self.drive_service = build('drive', 'v3', credentials=self.creds)
        self.gc = gspread.authorize(self.creds)
        self.folder_id = folder_id

    def list_files_in_folder(self) -> pd.DataFrame:
        """
        Lists all files in the specified Google Drive folder and returns a DataFrame
        with file names and Drive IDs.

        Returns:
            pd.DataFrame: A DataFrame containing 'Name' and 'Drive ID' columns for the files.
        """
        query = f"'{self.folder_id}' in parents"
        results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])

        file_data = {'Name': [], 'Drive ID': []}
        for file in files:
            file_data['Name'].append(file['name'])
            file_data['Drive ID'].append(file['id'])

        return pd.DataFrame(file_data)


    def get_drive_link(self, file_name: str, subfolder_id: str) -> str:
        """
        Retrieves the public Google Drive link for a file based on its name and subfolder ID.

        Args:
            file_name (str): The name of the file to search for in the Drive folder.
            subfolder_id (str): The Google Drive folder ID where the file is located.

        Returns:
            str: The Google Drive link to the file, or None if the file is not found.
        """
        query = f"name='{file_name}' and '{subfolder_id}' in parents"
        results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        items = results.get('files', [])
        
        if not items:
            return None
        else:
            file_id = items[0]['id']
            
            try:
                self.drive_service.permissions().create(
                    fileId=file_id,
                    body={'type': 'anyone', 'role': 'reader'},
                ).execute()
            except HttpError as error:
                print(f"Error setting permissions for file {file_name}: {error}")
                return None

            return f"https://drive.google.com/uc?id={file_id}"



    def list_subfolders_in_folder(self, parent_folder_id: str) -> dict:
        """
        Lists all subfolders in the specified Google Drive folder and returns a dictionary 
        with the folder names and their corresponding Drive IDs.

        Args:
            parent_folder_id (str): The Google Drive folder ID where to search for subfolders.

        Returns:
            dict: A dictionary with folder names as keys and their Drive IDs as values.
        """
        query = f"'{parent_folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder'"
        results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        folders = results.get('files', [])

        folder_data = {}
        for folder in folders:
            folder_data[folder['name']] = folder['id']

        return folder_data


    def create_sheet_with_data(self, spreadsheet_id: str, sheet_name: str, df: pd.DataFrame, image_cols: dict):
        """
        Uploads a DataFrame to a Google Sheet. If image columns are provided, it will attempt to 
        add images to the corresponding cells using Google Drive links from specified subfolders.

        Args:
            spreadsheet_id (str): The ID of the Google Spreadsheet where the data will be uploaded.
            sheet_name (str): The name of the sheet where data will be uploaded or created.
            df (pd.DataFrame): The DataFrame containing data to upload.
            image_cols (dict): A dictionary where keys are column names and values are corresponding folder IDs.
        """

        sh = self.gc.open_by_key(spreadsheet_id)
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=sheet_name, rows=str(len(df) + 10), cols=str(len(df.columns) + 10))

        worksheet.update([df.columns.values.tolist()] + df.fillna("").values.tolist())  

        if image_cols:
            for image_col, subfolder_id in image_cols.items():  
                for index, row in df.iterrows():
                    file_path = row[image_col]
                    
                    if pd.isna(file_path) or file_path == "":
                        worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, "no image")  
                        continue
                    
                    file_name_only = os.path.basename(file_path) 
                    drive_link = self.get_drive_link(file_name_only, subfolder_id)
                    
                    if drive_link:
                        image_formula = f'=IMAGE("{drive_link}")'
                        worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, image_formula)
                    else:
                        worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, "not found in drive")


    def add_column_with_drive_files(self, spreadsheet_id: str, sheet_name: str, df_files: pd.DataFrame, column_name: str, start_position: int):
        """
        Adds a new column to an existing Google Sheet starting from a specified position with Drive file links.

        Args:
            spreadsheet_id (str): The ID of the Google Spreadsheet where the column will be added.
            sheet_name (str): The name of the sheet where the column will be added.
            df_files (pd.DataFrame): The DataFrame containing file names and their Drive IDs.
            column_name (str): The name of the new column.
            start_position (int): The starting position (column index) where the new column will be added.
                                (1-based index, so 1 refers to column 'A').
        """
        
        sh = self.gc.open_by_key(spreadsheet_id)
        
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            raise Exception(f"Worksheet '{sheet_name}' not found in the spreadsheet.")
        
        worksheet.update_cell(1, start_position, column_name)
        
        for i, row in df_files.iterrows():
            file_link = f"https://drive.google.com/uc?id={row['Drive ID']}"
            worksheet.update_cell(i + 2, start_position, file_link) 

