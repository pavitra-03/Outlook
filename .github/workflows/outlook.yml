import os
import sys
import time
import requests
from datetime import datetime

# ============================================================
# CONFIGURATION
# ============================================================

TENANT_ID = os.getenv("MS_TENANT_ID")
CLIENT_ID = os.getenv("MS_CLIENT_ID")
CLIENT_SECRET = os.getenv("MS_CLIENT_SECRET")

MAILBOX = os.getenv("MS_MAILBOX")

# Number of emails to generate.
# Can be overridden from command line:
# py generate_message_trace_logs.py 1000
COUNT = int(sys.argv[1]) if len(sys.argv) > 1 else 30

# Keep below Exchange Online's documented 30 messages/minute
# per-user message rate.
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

token_url = (
    f"https://login.microsoftonline.com/"
    f"{TENANT_ID}/oauth2/v2.0/token"
)

token_data = {
    "client_id": CLIENT_ID,
    "client_secret": CLIENT_SECRET,
    "scope": "https://graph.microsoft.com/.default",
    "grant_type": "client_credentials",
}

print("Requesting Microsoft Graph access token...")

token_response = requests.post(
    token_url,
    data=token_data,
    timeout=30
)

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

send_url = (
    f"https://graph.microsoft.com/v1.0/"
    f"users/{MAILBOX}/sendMail"
)

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json",
}

print()
print("==========================================")
print("Message Trace Log Generator")
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

    subject = f"{SUBJECT_PREFIX} {i:04d}"

    body = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": (
                    f"Automated Microsoft Graph Message Trace "
                    f"test event #{i}.\n"
                    f"Generated at: {datetime.now().isoformat()}"
                ),
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": RECIPIENT
                    }
                }
            ],
        }
    }

    try:

        response = requests.post(
            send_url,
            headers=headers,
            json=body,
            timeout=30
        )

        if response.status_code == 202:
            successful += 1
            print(
                f"[{i}/{COUNT}] "
                f"202 Accepted - {subject}"
            )

        elif response.status_code == 429:
            print(
                f"[{i}/{COUNT}] "
                f"429 Throttled - waiting..."
            )

            retry_after = response.headers.get(
                "Retry-After",
                "60"
            )

            try:
                wait_seconds = int(retry_after)
            except ValueError:
                wait_seconds = 60

            print(
                f"Waiting {wait_seconds} seconds "
                f"before retry..."
            )

            time.sleep(wait_seconds)

            # Retry same message
            response = requests.post(
                send_url,
                headers=headers,
                json=body,
                timeout=30
            )

            if response.status_code == 202:
                successful += 1
                print(
                    f"[{i}/{COUNT}] "
                    f"202 Accepted after retry - {subject}"
                )
            else:
                failed += 1
                print(
                    f"[{i}/{COUNT}] "
                    f"Retry failed: "
                    f"{response.status_code}"
                )
                print(response.text)

        else:
            failed += 1

            print(
                f"[{i}/{COUNT}] "
                f"FAILED - HTTP {response.status_code}"
            )

            print(response.text)

    except requests.RequestException as e:

        failed += 1

        print(
            f"[{i}/{COUNT}] "
            f"REQUEST ERROR: {e}"
        )

    # Rate limiting
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
