#!/usr/bin/env python3
"""
Google Drive helper for the TCN report automation.
Handles OAuth login and listing / downloading / uploading workbooks.
"""

import io
import os

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

import config

SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
FOLDER_MIME = "application/vnd.google-apps.folder"


def get_service():
    """Authenticate (browser pops up the first time) and return a Drive client."""
    creds = None
    if os.path.exists(config.TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(config.TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(config.CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(config.TOKEN_FILE, "w") as fh:
            fh.write(creds.to_json())
    return build("drive", "v3", credentials=creds)


def list_files(service, folder_id):
    """All non-trashed files directly inside a folder."""
    files, page = [], None
    q = f"'{folder_id}' in parents and trashed = false"
    while True:
        resp = service.files().list(
            q=q, spaces="drive",
            fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
            pageToken=page,
        ).execute()
        files += resp.get("files", [])
        page = resp.get("nextPageToken")
        if not page:
            break
    return files


def list_subfolders(service, parent_id):
    """All non-trashed subfolders directly inside a parent folder."""
    folders, page = [], None
    q = (f"'{parent_id}' in parents and trashed = false "
         f"and mimeType = '{FOLDER_MIME}'")
    while True:
        resp = service.files().list(
            q=q, spaces="drive",
            fields="nextPageToken, files(id, name)",
            pageToken=page,
        ).execute()
        folders += resp.get("files", [])
        page = resp.get("nextPageToken")
        if not page:
            break
    return folders


def download_xlsx(service, file_obj, dest_path):
    """Download a file as .xlsx. Native Google Sheets are exported; real xlsx
    uploads are fetched as-is."""
    fid, mime = file_obj["id"], file_obj["mimeType"]
    if mime == GSHEET_MIME:
        request = service.files().export_media(fileId=fid, mimeType=XLSX_MIME)
    else:
        request = service.files().get_media(fileId=fid)
    with io.FileIO(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path


def upload_xlsx(service, local_path, folder_id, name):
    """Upload (or overwrite if same name already exists) an .xlsx into a folder."""
    existing = [f for f in list_files(service, folder_id)
                if f["name"] == name and f["mimeType"] != GSHEET_MIME]
    media = MediaFileUpload(local_path, mimetype=XLSX_MIME, resumable=True)
    if existing:
        return service.files().update(fileId=existing[0]["id"], media_body=media,
                                       fields="id, name, webViewLink").execute()
    meta = {"name": name, "parents": [folder_id]}
    return service.files().create(body=meta, media_body=media,
                                  fields="id, name, webViewLink").execute()
