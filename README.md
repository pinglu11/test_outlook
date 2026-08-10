# test_outlook

This repository contains a small script `outlook_mail.py` that uses the Microsoft Graph API to list recent Outlook messages using the device code flow.

## Quick start

1. Create and activate a virtual environment (recommended):

- macOS / Linux:

  ```bash
  python3 -m venv venv
  source venv/bin/activate
  ```

- Windows (PowerShell):

  ```ps1
  python -m venv venv
  .\venv\Scripts\Activate.ps1
  ```

2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Register an app in Azure (Microsoft Entra):

- Go to the Azure Portal → Azure Active Directory → App registrations → New registration.
- Name: e.g. `Outlook Device App`.
- Supported account types: choose based on your scenario (Accounts in this org / personal / etc).
- Register.
- In the app's Authentication settings, enable "Allow public client flows" (device code flow) for Public client applications.
- In API permissions → Microsoft Graph → Delegated permissions, add `Mail.Read` and grant admin consent if required.
- Copy the **Application (client) ID** — this is your OUTLOOK_CLIENT_ID.

4. Set environment variables:

- macOS / Linux:

  ```bash
  export OUTLOOK_CLIENT_ID="your-client-id"
  # optional: export OUTLOOK_TENANT_ID="your-tenant-id"
  ```

- Windows PowerShell:

  ```ps1
  $env:OUTLOOK_CLIENT_ID = "your-client-id"
  $env:OUTLOOK_TENANT_ID = "your-tenant-id" # optional
  ```

5. Run the script:

```bash
python outlook_mail.py
```

The script will print a verification URL and a user code. Open the URL in a browser and enter the code to sign in and consent. After successful authentication it will display recent messages.

## Troubleshooting

- If you see "Set OUTLOOK_CLIENT_ID...", ensure the environment variable is set in the same shell where you run the script.
- On headless machines, copy the verification URL to a browser on another device and enter the code.
- If you get permission errors, grant admin consent or sign in with an account that can consent to the requested permission.
