from __future__ import print_function
import os.path
import logging
import sys
import io
import os
from datetime import datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from pathlib import Path



# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive']

# Dont worry, I'm not leaking anything, these arent secrets, just the folder ID's
# The ID of the template google document.
TEMPLATE_DOCUMENT_ID = '1LjXMKG4M9j_uA-uCdvqdxVYs6Un0zIb0VPxKq2LQ9I0'
# The ID of the template google drive folder.
TEMPLATE_FOLDER_ID = '14tlAdhsmUXA0LbUE-l-WcbYyPKOPyTgL'
# The ID of the destination google drive storage folder
STORAGE_FOLDER_ID = '14tJoA7EpHmkmN73t6liTQ_meNO-k-7hb'


# Setting up logging
logger = logging.getLogger(__name__)
def setup_logging():
    FORMAT = '%(levelname)s:%(name)s: %(message)s (%(asctime)s; %(filename)s:%(lineno)d)'
    DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
    LEVEL = logging.INFO
    STREAM = sys.stdout
    logging.basicConfig( 
        level=LEVEL, 
        format=FORMAT, 
        datefmt=DATE_FORMAT,
        stream=STREAM,
    )


def get_tokem():
    """
    Verify and get access token for google docs and drive API's
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens
    # Automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


def get_current_date():
    """
    Get date in format: "21st December, 2023"
    """
    current_date = datetime.now()
    month_names = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ]

    day = current_date.day
    month = month_names[current_date.month - 1]
    year = current_date.year

    def get_day_suffix(day):
        if 10 <= day % 100 <= 20:
            suffix = "th"
        else:
            suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        return suffix

    day_suffix = get_day_suffix(day)

    # Create the formatted date string
    formatted_date = f"{day}{day_suffix} {month}, {year}"
    return formatted_date


def get_cv_data():
    """
    Prompt and collect position information
    """
    company_name = input('Company Name: ')
    company_location = input('Company Location: ')
    position = input('Position: ')
    position_team = input('Position Team Name: ')

    info = {
        'company' : company_name,
        'company_location': company_location,
        'position' : position,
        'position_team': position_team,
        'date': get_current_date(),
    }

    return info


def create_new_doc(drive_services, new_filename):
    """
    Cloning a copy of the CV tempate on Google Drive
    """
    try:
        # Cloning a new doc for cv template
        copy_request = {
            'name': new_filename, # Setting the new file name
            'parents': [STORAGE_FOLDER_ID],
        }

        # Copying the template cv file into the drive storage directory
        new_cv_copy = drive_services.files().copy(fileId=TEMPLATE_DOCUMENT_ID, body=copy_request).execute()
        return new_cv_copy

    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        logger.exception(f'Failed to generate new document copy in Google Drive')
        logger.error(error)
        return None


def stream_pdf_file(drive_services, doc_id, company_name, file_type="PDF"):
    """
    Download the file from Google Drive Services
    """
    try:
        request = drive_services.files().export(fileId=doc_id, mimeType=f'application/{file_type}')
        downloaded_file = io.BytesIO()
        downloader = MediaIoBaseDownload(downloaded_file, request)
        done = False

        while not done:
            status, done = downloader.next_chunk()
        
        # Save the downloaded content to a file
        # INFO: Currently saving the file in Downlaods folder locally
        with open(f'{str(Path.home() / "Downloads/")}/{company_name}_Cover_Letter_Aditya_Patel.{file_type}', 'wb') as output_file:
            downloaded_file.seek(0)
            output_file.write(downloaded_file.read())
    except Exception as e:
        logger.exception("Filed to download the file")
        logger.error(e)
        

def main():
    #INFO: Setting up the Google token
    creds = get_tokem()

    #INFO: Connecting to the Google Drive and Documents services
    try:
        drive_services = build('drive', 'v3', credentials=creds)
        document_services = build('docs', 'v1', credentials=creds)

    except Exception as e:
        logger.exception("❌ Unable to connect to the Google Drive Developer services")
        logger.error(e)
        exit(1)
    
    # Prompt Position Information
    populate_info = get_cv_data()
    logger.info("✨ Creating CV....")
    
    # Creating a new Doc copy from the template
    cv_file_copy = create_new_doc(drive_services, 'testing_filename')
    logger.info(f"📄 New file copy created. ID: {cv_file_copy}")
    logger.info(f"💻 Populating the document.....")

    add_cv_info = [{'replaceAllText': {'containsText': {'text': "{{POSITION}}", 'matchCase': True}, 'replaceText': populate_info['position']}},
                     {'replaceAllText': {'containsText': {'text': "{{COMPANY}}", 'matchCase': True}, 'replaceText': populate_info['company']}},
                     {'replaceAllText': {'containsText': {'text': "{{COMPANY_LOCATION}}", 'matchCase': True}, 'replaceText': populate_info['company_location']}},
                     {'replaceAllText': {'containsText': {'text': "{{DATE}}", 'matchCase': True}, 'replaceText': populate_info['date']}},
                     {'replaceAllText': {'containsText': {'text': "{{POSITION_TEAM}}", 'matchCase': True}, 'replaceText': populate_info['position_team']}}]

    # Poppulating the CV with information
    document_services.documents().batchUpdate(documentId=cv_file_copy['id'], body={'requests': add_cv_info}).execute()

    # Download the PDF file
    logger.info("🤠 Downloading the PDF....")
    stream_pdf_file(drive_services, cv_file_copy['id'], populate_info['company'].replace(' ', '_'))

    # Deleting the document from google drive
    logger.info("🗑️ Deleting the Google Document....")
    try:
        drive_services.files().delete(fileId=cv_file_copy['id']).execute()
    except Exception as e:
        logger.exception("Failed to delete the Google Document")
        logger.error(e)
        exit(1)
    logger.info("✅ CV Downloaded in your DOWNLOADS folder")

if __name__ == '__main__':
    setup_logging()
    main()