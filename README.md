# SharePoint Document Data Cleaning Tool

A Python tool that connects to Microsoft OneDrive/SharePoint via the Microsoft Graph API, recursively scans all folders, downloads documents (PDF, DOCX), and extracts their text content for data cleaning and processing.

## Project Structure

```
AI-Data-Cleaning-Tool/
├── auth.py              # Microsoft Azure AD authentication (device code flow)
├── config.py            # Configuration loader (reads from .env)
├── sharepoint_api.py    # OneDrive/SharePoint file listing & downloading (recursive)
├── text_extraction.py   # Text extraction from PDF and DOCX files
├── main.py              # Full pipeline: authenticate -> list -> download -> extract
├── test_setup.py        # Quick verification that everything is installed
├── setup.sh             # Automated setup script
├── requirements.txt     # Python dependencies
├── .env.example         # Template for environment variables
├── downloaded_files/    # Downloaded documents end up here
├── extracted_text/      # Extracted text files end up here
└── logs/                # Log files and summary reports
```

## How It Works

1. **Authenticate** with Microsoft using device code flow (no web server needed)
2. **Recursively scan** all OneDrive/SharePoint folders for `.docx` and `.pdf` files
3. **Download** matching files locally
4. **Extract text** from each file (supports Word and PDF)
5. **Save** extracted text and a summary report

## Prerequisites

- Python 3.9+
- A Microsoft account (personal or work/school)
- An Azure AD App Registration

## Setup

### 1. Clone and install dependencies

```bash
git clone <repo-url>
cd AI-Data-Cleaning-Tool

# Option A: Use the setup script
chmod +x setup.sh
./setup.sh

# Option B: Manual setup
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 2. Create an Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Azure Active Directory** > **App registrations** > **New registration**
2. Name: whatever you want (e.g. "Data Cleaning Tool")
3. Supported account types: **"Accounts in any organizational directory and personal Microsoft accounts"**
4. Click **Register**
5. Copy your **Application (client) ID** and **Directory (tenant) ID** from the Overview page

### 3. Configure API permissions

1. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
2. Add: `Files.Read` (for personal accounts) or `Files.Read.All` (for SharePoint)
3. Add: `User.Read`

### 4. Enable device code flow

1. Go to **Authentication** > **Advanced settings**
2. Set **Allow public client flows** to **Yes**
3. Under **Mobile and desktop applications**, add redirect URIs:
   - `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - `http://localhost`

### 5. Configure environment variables

```bash
cp .env.example .env
```

Edit `.env`:

```env
CLIENT_ID=your-client-id
TENANT_ID=your-tenant-id

# For personal Microsoft accounts (Outlook/Hotmail):
AUTHORITY=https://login.microsoftonline.com/consumers
SCOPES=Files.Read User.Read

# For work/school accounts (SharePoint):
# AUTHORITY=https://login.microsoftonline.com/your-tenant-id
# SCOPES=Files.Read.All User.Read
```

## Usage

### Verify setup

```bash
python test_setup.py
```

### Authenticate

```bash
python auth.py
```

1. A code appears in the terminal (e.g. `L3K965WG3`)
2. Open https://microsoft.com/devicelogin in your browser
3. Enter the code and sign in with your Microsoft account
4. Approve the permissions
5. Close the browser tab (the localhost error is normal)
6. Check the terminal for success message

Token is cached locally so you only authenticate once until it expires.

### List all files

```bash
python sharepoint_api.py
```

Recursively scans all folders and lists every PDF and Word file, grouped by type with folder paths.

### Run full pipeline

```bash
python main.py
```

Downloads and extracts text from all matching files. Currently limited to 5 files (change `limit=5` in `main.py` to `limit=None` for all files).

## Authority Endpoints

| Account Type | Authority URL |
|---|---|
| Personal (Outlook/Hotmail) | `https://login.microsoftonline.com/consumers` |
| Work/School only | `https://login.microsoftonline.com/{tenant-id}` |
| Both | `https://login.microsoftonline.com/common` |

## Troubleshooting

| Problem | Solution |
|---|---|
| "This site can't be reached" (localhost) after browser sign-in | Normal for device code flow. Close the tab, check your terminal. |
| No consent screen in browser | Change `AUTHORITY` to `/consumers` for personal accounts |
| `Files.Read.All` errors with personal account | Use `Files.Read` instead (without `.All`) |
| Token expired | Delete `.msal_token_cache.json` and run `python auth.py` again |
| Only finding root-level files | Make sure you're using the latest `sharepoint_api.py` with recursive scanning |
