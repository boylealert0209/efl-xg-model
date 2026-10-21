"""
upload_to_drive.py

Replace the contents of an EXISTING Google Sheet with a freshly built .xlsx
report, so the Sheet (and any link/bookmark to it) updates automatically after
each scheduled run. Drive converts the .xlsx on upload, so formulas, colours
and charts come across.

One-off setup (see the chat explanation): create a Google Cloud service
account, put its JSON key in the GitHub secret GDRIVE_SA_JSON, and share the
Google Sheet with the service account's email address as an Editor.

    python upload_to_drive.py --file data/championship_xg_report.xlsx --sheet-id <ID>

<ID> is the long string in the Sheet's URL: docs.google.com/spreadsheets/d/<ID>/edit
"""
import argparse
import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--sheet-id", required=True)
    args = ap.parse_args()

    key = os.environ.get("GDRIVE_SA_JSON")
    if not key:
        sys.exit("GDRIVE_SA_JSON is not set")
    creds = service_account.Credentials.from_service_account_info(
        json.loads(key), scopes=["https://www.googleapis.com/auth/drive"])
    drive = build("drive", "v3", credentials=creds, cache_discovery=False)
    media = MediaFileUpload(args.file, mimetype=XLSX, resumable=False)
    drive.files().update(fileId=args.sheet_id, media_body=media,
                         supportsAllDrives=True).execute()
    print(f"Updated Google Sheet {args.sheet_id} from {args.file}")


if __name__ == "__main__":
    main()
