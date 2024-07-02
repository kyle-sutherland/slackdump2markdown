import os
import json
import argparse
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.http import MediaFileUpload


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

    document = (
        docs_service.documents()
        .create(body={"title": "Slack Conversation Log"})
        .execute()
    )
    document_id = document["documentId"]

    requests = [
        {"insertText": {"location": {"index": 1}, "text": "Slack Conversation Log\n"}},
        {
            "updateParagraphStyle": {
                "range": {"startIndex": 1, "endIndex": 24},
                "paragraphStyle": {"namedStyleType": "HEADING_1"},
                "fields": "namedStyleType",
            }
        },
        {
            "updateTextStyle": {
                "range": {"startIndex": 1, "endIndex": 2},
                "textStyle": {"fontSize": {"magnitude": 11, "unit": "PT"}},
                "fields": "fontSize",
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
                file_metadata = {"name": attachment["name"]}
                media = MediaFileUpload(file_path, resumable=True)
                file = (
                    drive_service.files()
                    .create(body=file_metadata, media_body=media, fields="id")
                    .execute()
                )

                requests.extend(
                    [
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
                                "uri": f"https://drive.google.com/uc?id={file['id']}",
                                "objectSize": {
                                    "height": {"magnitude": 200, "unit": "PT"},
                                    "width": {"magnitude": 200, "unit": "PT"},
                                },
                            }
                        },
                    ]
                )

    docs_service.documents().batchUpdate(
        documentId=document_id, body={"requests": requests}
    ).execute()
    print(f"Document created: https://docs.google.com/document/d/{document_id}/edit")


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
