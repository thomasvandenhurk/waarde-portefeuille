from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload


def copy_to_gsheet(excel_output, folder_id):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name('keyfile.json', scope)
    file_name = excel_output.split("\\")[-1].replace('.xlsx', '')

    # build the Google Drive API client
    drive_service = build('drive', 'v3', credentials=creds)

    existing_id = drive_service.files().list(
        q=f"name contains '{file_name}' and '{folder_id}' in parents",
        spaces='drive',
        fields='files(id, name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True
    ).execute()
    items = existing_id.get('files', [])

    media = MediaFileUpload(
        excel_output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

    if items:
        # overwrite existing file
        file_id = items[0]['id']
        drive_service.files().update(
            fileId=file_id,
            media_body=media
        ).execute()
    else:
        # create new file
        file_metadata = {
            'name': f'{file_name}',
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': [folder_id],
        }

        # Upload the file and convert to Google Sheets
        drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
