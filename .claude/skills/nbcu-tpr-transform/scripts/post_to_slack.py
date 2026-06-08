#!/usr/bin/env python3
"""Usage: python post_to_slack.py <file_path> <summary> [unknown_mpms_csv_path]"""
import os, sys
from slack_sdk import WebClient

client = WebClient(token=os.environ["SLACK_BOT_TOKEN"])
channel = os.environ["SLACK_CHANNEL_ID"]
file_path, summary = sys.argv[1], sys.argv[2]
unknown_mpms_csv = sys.argv[3] if len(sys.argv) > 3 and sys.argv[3] else None

with open(file_path, "rb") as f:
    client.files_upload_v2(
        channel=channel,
        file=f,
        filename=os.path.basename(file_path),
        initial_comment=summary,
    )

if unknown_mpms_csv:
    with open(unknown_mpms_csv, "rb") as f:
        client.files_upload_v2(
            channel=channel,
            file=f,
            filename="unknown_mpms.csv",
            title="Unknown MPMs",
            initial_comment="⚠️ These MPMs were not found in the reference title list:",
        )

print("Posted to Slack.")
