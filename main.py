import os
import json
import argparse
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError


def parse_json_files(directory):
    conversations = []
    for filename in os.listdir(directory):
        if filename.endswith(".json"):
            with open(os.path.join(directory, filename), "r") as file:
                data = json.load(file)
                for message in data:
                    timestamp = datetime.fromtimestamp(float(message["ts"]))
                    attachments = []
                    if "files" in message:
                        for file in message["files"]:
                            attachments.append(
                                {
                                    "name": file["name"],
                                    "url": file["url_private"],
                                    "local_path": os.path.join(
                                        "attachments",
                                        file["url_private"].split("/")[-1],
                                    ),
                                }
                            )
                    conversations.append(
                        {
                            "date": timestamp.strftime("%Y-%m-%d"),
                            "time": timestamp.strftime("%H:%M:%S"),
                            "user": message["user_profile"]["real_name"],
                            "text": message.get("text", ""),
                            "attachments": attachments,
                        }
                    )
    return sorted(conversations, key=lambda x: (x["date"], x["time"]))


def create_markdown(conversations, output_dir):
    markdown = "# Slack Conversation Log\n\n"
    for message in conversations:
        markdown += f"**[{message['date']} {message['time']}] - {message['user']}:**\n\n{message['text']}\n\n"
        for attachment in message["attachments"]:
            markdown += (
                f"[Attachment: {attachment['name']}]({attachment['local_path']})\n\n"
            )
        markdown += "----\n\n"

    output_file = os.path.join(output_dir, "output.md")
    with open(output_file, "w") as f:
        f.write(markdown)
    return markdown


def create_formatted_googledoc(conversations, directory):
    SCOPES = [
        "https://www.googleapis.com/auth/documents",
        "https://www.googleapis.com/auth/drive.file",
    ]
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # Create a new folder in Google Drive
    folder_name = os.path.basename(os.path.normpath(directory))
    folder_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
    }
    folder = drive_service.files().create(body=folder_metadata, fields="id").execute()
    folder_id = folder.get("id")

    # Create attachments subfolder
    attachments_folder_metadata = {
        "name": "attachments",
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [folder_id],
    }
    attachments_folder = (
        drive_service.files()
        .create(body=attachments_folder_metadata, fields="id")
        .execute()
    )
    attachments_folder_id = attachments_folder.get("id")

    # Create the document in the new folder
    document = (
        docs_service.documents()
        .create(body={"title": "Slack Conversation Log"})
        .execute()
    )
    document_id = document["documentId"]

    # Move the document to the new folder
    drive_service.files().update(
        fileId=document_id, addParents=folder_id, fields="id, parents"
    ).execute()

    requests = [
        {"insertText": {"location": {"index": 1}, "text": "Slack Conversation Log\n"}},
        {
            "updateParagraphStyle": {
                "range": {"startIndex": 1, "endIndex": 24},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType",
            }
        },
    ]

    for message in conversations:
        timestamp = f"{message['date']} {message['time']}"
        user = message["user"]
        text = message["text"]

        start_index = len(
            "".join(req.get("insertText", {}).get("text", "") for req in requests)
        )

        requests.extend(
            [
                {
                    "insertText": {
                        "location": {"index": start_index},
                        "text": f"{user} [{timestamp}]: {text}\n",
                    }
                },
                {
                    "updateTextStyle": {
                        "range": {
                            "startIndex": start_index,
                            "endIndex": start_index
                            + len(user)
                            + len(timestamp)
                            + len(text)
                            + 5,
                        },
                        "textStyle": {"fontSize": {"magnitude": 11, "unit": "PT"}},
                        "fields": "fontSize",
                    }
                },
                {
                    "updateTextStyle": {
                        "range": {
                            "startIndex": start_index,
                            "endIndex": start_index + len(user),
                        },
                        "textStyle": {"bold": True},
                        "fields": "bold",
                    }
                },
                {
                    "updateTextStyle": {
                        "range": {
                            "startIndex": start_index + len(user) + 1,
                            "endIndex": start_index + len(user) + len(timestamp) + 3,
                        },
                        "textStyle": {
                            "foregroundColor": {
                                "color": {
                                    "rgbColor": {"red": 0.5, "green": 0.5, "blue": 0.5}
                                }
                            }
                        },
                        "fields": "foregroundColor",
                    }
                },
            ]
        )

        for attachment in message["attachments"]:
            file_path = os.path.join(directory, attachment["local_path"])
            if os.path.exists(file_path):
                file_metadata = {
                    "name": attachment["name"],
                    "parents": [attachments_folder_id],
                }
                media = MediaFileUpload(file_path, resumable=True)
                file = (
                    drive_service.files()
                    .create(
                        body=file_metadata,
                        media_body=media,
                        fields="id, mimeType, webContentLink",
                    )
                    .execute()
                )

                mime_type = file.get("mimeType", "")

                requests.extend(
                    [
                        {
                            "updateTextStyle": {
                                "range": {
                                    "startIndex": start_index,
                                    "endIndex": start_index
                                    + len(user)
                                    + len(timestamp)
                                    + len(text)
                                    + 5,
                                },
                                "textStyle": {
                                    "fontSize": {"magnitude": 11, "unit": "PT"}
                                },
                                "fields": "fontSize",
                            },
                        },
                        {
                            "insertText": {
                                "location": {
                                    "index": start_index
                                    + len(user)
                                    + len(timestamp)
                                    + len(text)
                                    + 4
                                },
                                "text": f"\nAttachment: {attachment['name']}\n",
                            }
                        },
                    ]
                )

                if mime_type.startswith("image/"):
                    # Create a publicly accessible link
                    drive_service.permissions().create(
                        fileId=file["id"],
                        body={"type": "anyone", "role": "reader"},
                        fields="id",
                    ).execute()

                    web_content_link = file.get("webContentLink", "")
                    if web_content_link:
                        requests.append(
                            {
                                "insertInlineImage": {
                                    "location": {
                                        "index": start_index
                                        + len(user)
                                        + len(timestamp)
                                        + len(text)
                                        + 5
                                        + len(attachment["name"])
                                        + 12
                                    },
                                    "uri": web_content_link,
                                    "objectSize": {
                                        "height": {"magnitude": 200, "unit": "PT"},
                                        "width": {"magnitude": 200, "unit": "PT"},
                                    },
                                }
                            }
                        )
                    else:
                        requests.append(
                            {
                                "insertText": {
                                    "location": {
                                        "index": start_index
                                        + len(user)
                                        + len(timestamp)
                                        + len(text)
                                        + 5
                                        + len(attachment["name"])
                                        + 12
                                    },
                                    "text": f"Image could not be inserted. View it here: https://drive.google.com/file/d/{file['id']}/view\n",
                                }
                            }
                        )
                else:
                    # For non-image files, insert a link to the file
                    requests.append(
                        {
                            "insertText": {
                                "location": {
                                    "index": start_index
                                    + len(user)
                                    + len(timestamp)
                                    + len(text)
                                    + 5
                                    + len(attachment["name"])
                                    + 12
                                },
                                "text": f"Link to file: https://drive.google.com/file/d/{file['id']}/view\n",
                            }
                        }
                    )

    try:
        docs_service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()
        print(
            f"Document created: https://docs.google.com/document/d/{document_id}/edit"
        )
        print(f"Folder created: https://drive.google.com/drive/folders/{folder_id}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("Requests that caused the error:")
        for i, request in enumerate(requests):
            print(f"Request {i}:")
            print(json.dumps(request, indent=2))


def main():
    parser = argparse.ArgumentParser(
        description="Parse Slack JSON files and create a Google Doc"
    )
    parser.add_argument("directory", help="Directory containing JSON files")
    args = parser.parse_args()

    conversations = parse_json_files(args.directory)
    create_markdown(conversations, args.directory)
    create_formatted_googledoc(conversations, args.directory)


if __name__ == "__main__":
    main()
