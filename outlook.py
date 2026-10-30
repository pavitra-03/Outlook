import os
import sys
import time
import uuid
import requests
from datetime import datetime, timezone

# ============================================================
# CONFIGURATION
# ============================================================

TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")

MAILBOX = os.getenv("MS_MAILBOX")

# Number of emails to generate.
COUNT = int(sys.argv[1]) if len(sys.argv) > 1 else 30

# Keep below Exchange Online's documented 30 messages/minute per-user message rate.
MESSAGES_PER_MINUTE = 25

# Calculate delay between messages.
DELAY = 60 / MESSAGES_PER_MINUTE

SUBJECT_PREFIX = "Message Trace Automated Test"

# Send to the same mailbox by default.
RECIPIENT = os.getenv("MS_RECIPIENT") or MAILBOX

# ============================================================
# VALIDATION
# ============================================================

required = {
    "MS_TENANT_ID": TENANT_ID,
    "MS_CLIENT_ID": CLIENT_ID,
    "MS_CLIENT_SECRET": CLIENT_SECRET,
    "MS_MAILBOX": MAILBOX,
}

missing = [name for name, value in required.items() if not value]

if missing:
    print("Missing environment variables:")
    for item in missing:
        print(f"  {item}")
    sys.exit(1)

# ============================================================
# GET APPLICATION TOKEN
# ============================================================

token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"

token_data = {
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials",
}

print("Requesting Microsoft Graph access token...")

token_response = requests.post(token_url, data=token_data, timeout=30)

if token_response.status_code != 200:
    print("Failed to obtain access token.")
    print(token_response.status_code)
    print(token_response.text)
    sys.exit(1)

access_token = token_response.json()["access_token"]
print("Access token obtained successfully.")

# ============================================================
# GRAPH SENDMAIL
# ============================================================

send_url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/sendMail"

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

print()
print("==========================================")
print("Message Trace Log Generator (Unique Emails)")
print("==========================================")
print(f"Mailbox       : {MAILBOX}")
print(f"Recipient     : {RECIPIENT}")
print(f"Messages      : {COUNT}")
print(f"Rate          : {MESSAGES_PER_MINUTE}/minute")
print(f"Delay         : {DELAY:.2f} seconds")
print("==========================================")
print()

successful = 0
failed = 0

for i in range(1, COUNT + 1):
    # 1. Generate unique identifiers for each message
    unique_id = str(uuid.uuid4())
    timestamp_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + "Z"
    
    # 2. Build unique Subject line and Custom Message-ID Header
    subject = f"{SUBJECT_PREFIX} [{unique_id[:8]}] - #{i:04d}"
    custom_message_id = f"<{unique_id}@{MAILBOX.split('@')[-1]}>"

    body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": (
                    f"Automated Microsoft Graph Message Trace test event #{i}.\n"
                    f"Unique Message ID : {unique_id}\n"
                    f"Generated at UTC   : {timestamp_str}"
                ),
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": RECIPIENT
                    }
                }
            ],
            # Pass custom headers to enforce distinct InternetMessageId and tracking tags
            "internetMessageHeaders": [
                {
                    "name": "X-Custom-Message-ID",
                    "value": custom_message_id
                },
                {
                    "name": "X-Test-Run-UUID",
                    "value": unique_id
                }
            ]
        }
    }

    try:
        response = requests.post(send_url, headers=headers, json=body, timeout=30)

        if response.status_code == 202:
            successful += 1
            print(f"[{i}/{COUNT}] 202 Accepted | UUID: {unique_id[:8]} | {subject}")

        elif response.status_code == 429:
            print(f"[{i}/{COUNT}] 429 Throttled - waiting...")
            retry_after = response.headers.get("Retry-After", "60")
            try:
                wait_seconds = int(retry_after)
            except ValueError:
                wait_seconds = 60

            print(f"Waiting {wait_seconds} seconds before retry...")
            time.sleep(wait_seconds)

            # Retry sending the message
            response = requests.post(send_url, headers=headers, json=body, timeout=30)

            if response.status_code == 202:
                successful += 1
                print(f"[{i}/{COUNT}] 202 Accepted after retry - {subject}")
            else:
                failed += 1
                print(f"[{i}/{COUNT}] Retry failed: {response.status_code}\n{response.text}")

        else:
            failed += 1
            print(f"[{i}/{COUNT}] FAILED - HTTP {response.status_code}\n{response.text}")

    except requests.RequestException as e:
        failed += 1
        print(f"[{i}/{COUNT}] REQUEST ERROR: {e}")

    # Rate limiting sleep
    if i < COUNT:
        time.sleep(DELAY)

# ============================================================
# SUMMARY
# ============================================================

print()
print("==========================================")
print("GENERATION COMPLETE")
print("==========================================")
print(f"Requested : {COUNT}")
print(f"Successful: {successful}")
print(f"Failed    : {failed}")
print("==========================================")
