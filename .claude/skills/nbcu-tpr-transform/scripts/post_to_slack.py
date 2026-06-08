#!/usr/bin/env python3
"""Usage: python scripts/post_to_slack.py <file_path> <summary>"""
import os, sys
from slack_sdk import WebClient

client = WebClient(token=os.environ["SLACK_BOT_TOKEN"])
file_path, summary = sys.argv[1], sys.argv[2]

with open(file_path, "rb") as f:
    client.files_upload_v2(
        channel=os.environ["SLACK_CHANNEL_ID"],
        file=f,
        filename=os.path.basename(file_path),
        initial_comment=summary,
    )
print("Posted to Slack.")
