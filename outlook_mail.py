import os
import sys
import json
import webbrowser
from typing import List

import msal
import requests


CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID")
TENANT_ID = os.environ.get("OUTLOOK_TENANT_ID", "common")
SCOPES = ["Mail.Read"]
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_URL = "https://graph.microsoft.com/v1.0/me/messages"


def acquire_access_token() -> str:
    if not CLIENT_ID:
        raise SystemExit(
            "Set OUTLOOK_CLIENT_ID to your Microsoft Entra app registration client ID before running this script."
        )

    app = msal.PublicClientApplication(client_id=CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        raise RuntimeError(json.dumps(flow, indent=2))

    print("Sign in to Microsoft to access your Outlook mailbox.")
    print(f"Open: {flow['verification_uri']}")
    print(f"Enter code: {flow['user_code']}")
    webbrowser.open(flow["verification_uri"])

    token_response = app.acquire_token_by_device_flow(flow)
    if "access_token" not in token_response:
        raise RuntimeError(json.dumps(token_response, indent=2))

    return token_response["access_token"]


def list_messages(access_token: str, limit: int = 10) -> List[dict]:
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {
        "$top": limit,
        "$select": "subject,from,receivedDateTime,bodyPreview",
        "$orderby": "receivedDateTime desc",
    }

    response = requests.get(GRAPH_URL, headers=headers, params=params, timeout=30)
    response.raise_for_status()
    payload = response.json()
    return payload.get("value", [])


def main() -> None:
    access_token = acquire_access_token()
    messages = list_messages(access_token, limit=10)

    if not messages:
        print("No messages found.")
        return

    print("\nRecent messages:")
    for message in messages:
        from_name = message.get("from", {}).get("emailAddress", {}).get("name", "Unknown")
        from_email = message.get("from", {}).get("emailAddress", {}).get("address", "")
        subject = message.get("subject") or "(no subject)"
        received = message.get("receivedDateTime", "")
        preview = (message.get("bodyPreview") or "").replace("\n", " ")
        print(f"- {received} | {from_name} <{from_email}> | {subject}")
        if preview:
            print(f"  {preview}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
