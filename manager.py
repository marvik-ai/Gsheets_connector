import os
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
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
    
    def __init__(self, service_account_file: str = None, folder_id: str = None, use_dotenv: bool = False):
        """
        Initializes the GoogleDriveManager with Google API credentials and Drive folder ID.

        Args:
            service_account_file (str, optional): Path to the Google Service Account JSON file.
                If not provided, it will look for the credentials in an environment variable or .env file.
            folder_id (str): Google Drive folder ID for file operations.
            use_dotenv (bool, optional): If True, loads credentials from a .env file. Default is False.
        """
        SCOPES = [
            'https://www.googleapis.com/auth/drive',
            'https://www.googleapis.com/auth/spreadsheets'
        ]
        
        if use_dotenv:
            load_dotenv()
            
        service_account_info = os.getenv('GOOGLE_CREDENTIALS_JSON')
        
        if service_account_info:
            creds_dict = json.loads(service_account_info)
            self.creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        elif service_account_file:
            self.creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
        else:
            raise ValueError("No se proporcionó un archivo de credenciales, variable de entorno ni .env.")

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

    def get_drive_link(self, file_name: str) -> str:
        """
        Retrieves the public Google Drive link for a file based on its name.

        Args:
            file_name (str): The name of the file to search for in the Drive folder.

        Returns:
            str: The Google Drive link to the file, or None if the file is not found.
        """
        query = f"name='{file_name}' and '{self.folder_id}' in parents"
        results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        items = results.get('files', [])
        if not items:
            return None
        else:
            file_id = items[0]['id']
            
            self.drive_service.permissions().create(
                fileId=file_id,
                body={'type': 'anyone', 'role': 'reader'},
            ).execute()  #Verifier

            return f"https://drive.google.com/uc?id={file_id}"


    def create_sheet_with_data(self, spreadsheet_id: str, sheet_name: str, df: pd.DataFrame, image_col: str = None):
        """
        Uploads a DataFrame to a Google Sheet. If an image column is provided, it will attempt to 
        add images to the corresponding cells using Google Drive links.

        Args:
            spreadsheet_id (str): The ID of the Google Spreadsheet where the data will be uploaded.
            sheet_name (str): The name of the sheet where data will be uploaded or created.
            df (pd.DataFrame): The DataFrame containing data to upload.
            image_col (str, optional): The name of the column that contains image file names. If provided,
                the method will attempt to locate the images in Google Drive and insert them in the sheet.
        """
        
        sh = self.gc.open_by_key(spreadsheet_id)
        try:
            worksheet = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            worksheet = sh.add_worksheet(title=sheet_name, rows=str(len(df) + 10), cols=str(len(df.columns) + 10))

        worksheet.update([df.columns.values.tolist()] + df.fillna("").values.tolist()) 

        if image_col:
            for index, row in df.iterrows():
                file_name = row[image_col]
                
                if pd.isna(file_name) or file_name == "":
                    worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, "no image")
                    continue
                
                file_name_only = os.path.basename(file_name) 
                drive_link = self.get_drive_link(file_name_only)
                
                if drive_link:
                    image_formula = f'=IMAGE("{drive_link}")'
                    worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, image_formula)
                else:
                    worksheet.update_cell(index + 2, df.columns.get_loc(image_col) + 1, "not found in drive")
