# Python Outlook Graph API CLI

A command-line interface for managing Microsoft Outlook emails using the Microsoft Graph API.

## Features

- Silent authentication with Windows Account Manager (WAM) and token caching
- Read, list, move, and delete emails
- Manage mail folders
- Full email body access
- Built with async/await for efficient API calls
- Pipe-delimited output format for easy scripting and AI integration
- Background data loading for improved performance

## Prerequisites

- Python 3.13+
- Microsoft Azure AD app registration
- Microsoft 365 / Outlook account

## Setup

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Create Azure AD App Registration

1. Go to [Microsoft Entra admin center](https://entra.microsoft.com) → Applications → App registrations
2. Create a new registration:
   - **Name**: Python Outlook CLI (or your preferred name)
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - **Redirect URI**: Leave blank for now

3. After creation, note the **Application (client) ID** and **Directory (tenant) ID**

4. Configure **API Permissions**:
   - Add Microsoft Graph permissions:
     - `User.Read`
     - `Mail.ReadWrite`
     - `Mail.Send`
   - Grant admin consent if required

5. Configure **Authentication**:
   - Go to Authentication → Add a platform → Mobile and desktop applications
   - Add these redirect URIs:
     - `http://localhost:8400`
     - `ms-appx-web://microsoft.aad.brokerplugin/{YOUR_CLIENT_ID}`
       (replace `{YOUR_CLIENT_ID}` with your actual Application ID)

### 3. Create Configuration File

Create a file named `.auth.json` in the project root:

```json
{
  "clientId": "YOUR_APPLICATION_CLIENT_ID",
  "tenantId": "YOUR_DIRECTORY_TENANT_ID",
  "graphUserScopes": "User.Read Mail.ReadWrite Mail.Send"
}
```

Replace:
- `YOUR_APPLICATION_CLIENT_ID` with your Application (client) ID from Azure
- `YOUR_DIRECTORY_TENANT_ID` with your Directory (tenant) ID from Azure

**Note**: `.auth.json` is excluded from version control for security.

### 4. Authenticate

Run the login command to authenticate:

```bash
python -m outlook login
```

This will:
- Open a browser window for authentication
- Save your authentication record to `.auth_record.json` (auto-generated, excluded from version control)
- Enable token caching for silent authentication on subsequent runs

## Usage

All commands are run using `python -m outlook <command>`.

### Authentication

#### `login`
Authenticate with your Microsoft account and save credentials.

```bash
python -m outlook login
```

Output:
```
OK|authenticated
```

### User Information

#### `user`
Display current user information (pipe-delimited format).

```bash
python -m outlook user
```

Output:
```
John Doe|john.doe@example.com
```

### Managing Folders

#### `folders`
List all mail folders with their IDs and metadata (pipe-delimited format).

```bash
python -m outlook folders
```

Output:
```
Inbox|AAMkAD...|parent=NONE|children=5|total=342|unread=12|hidden=false
Archive|AAMkAD...|parent=AAMkAD...|children=0|total=128|unread=0|hidden=false
Deleted Items|AAMkAD...|parent=NONE|children=0|total=45|unread=2|hidden=false
```

### Listing Messages

#### `list [FOLDER_ID] [--top N] [--oldest-first]`
List messages from a specific folder (pipe-delimited format with full metadata).

```bash
# List inbox (default, 100 messages, newest first)
python -m outlook list

# List specific folder
python -m outlook list AAMkAD...

# List 50 messages
python -m outlook list inbox --top 50
python -m outlook list --top 50  # Short form with default inbox

# List oldest messages first (useful for incremental cleanup)
python -m outlook list inbox --oldest-first
python -m outlook list inbox --top 50 --oldest-first
```

Output includes (pipe-delimited):
- Message ID
- Subject
- From (name and email)
- To recipients
- CC recipients
- Read/Unread status
- Received date/time
- Sent date/time
- Has attachments
- Importance level
- Conversation ID
- Parent folder ID
- Web link

### Reading Messages

#### `read MESSAGE_ID`
Display full message including body content (pipe-delimited header + body).

```bash
python -m outlook read AAMkAD...
```

Output:
```
id=AAMkAD...|subject=Meeting Tomorrow|from=Jane Smith <jane@example.com>|to=...|cc=...|received=2025-10-01 14:26:09+00:00|sent=2025-10-01 14:25:00+00:00|status=read|attachments=false|importance=normal|conversation=AAMkAD...|folder=AAMkAD...|weblink=https://...

--- Body (html) ---
<html>Hi team,<br>Let's meet tomorrow at 2pm...</html>
```

### Moving Messages

#### `move MESSAGE_ID DESTINATION_FOLDER_ID`
Move a message to a different folder.

```bash
python -m outlook move AAMkAD... AAMkAD...
```

Output:
```
OK|moved|AAMkAD...|to|AAMkAD...
```

### Deleting Messages

#### `delete MESSAGE_ID`
Move a message to Deleted Items folder (soft delete).

```bash
python -m outlook delete AAMkAD...
```

Output:
```
OK|deleted|AAMkAD...
```

**Note**: This performs a soft delete (moves to Deleted Items). Permanent deletion is not available via the CLI.

## Configuration Files

### `.auth.json` (Required - You Create)
Contains your Azure AD app credentials. Must be created manually.

```json
{
  "clientId": "...",
  "tenantId": "...",
  "graphUserScopes": "User.Read Mail.ReadWrite Mail.Send"
}
```

### `.auth_record.json` (Auto-generated)
Contains your authentication record. Created automatically when you run `login` command.
- Stores account information for silent authentication
- Excluded from version control
- Delete this file to force re-authentication

## Development

### Running from Virtual Environment

```bash
# Activate virtual environment (Linux/WSL)
source venv/bin/activate

# Or on Windows
venv\Scripts\activate

# Run commands
python -m outlook <command>
```

### Running Without Activating Virtual Environment

```bash
# Linux/WSL
./venv/bin/python -m outlook <command>

# Windows
venv\Scripts\python.exe -m outlook <command>
```

## Troubleshooting

### Authentication Issues

If you encounter authentication errors:

1. Delete `.auth_record.json`:
   ```bash
   rm .auth_record.json
   ```

2. Run login again:
   ```bash
   python -m outlook login
   ```

### Permission Errors

If you get permission errors:

1. Verify your Azure AD app has the correct permissions:
   - `User.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`

2. Ensure admin consent has been granted

3. Verify redirect URIs are configured correctly

### Browser Not Opening (WSL)

If you're using WSL and the browser doesn't open automatically, manually navigate to the URL shown in the terminal.

## Architecture

- **`outlook/__main__.py`**: CLI commands using Click with async support and pipe-delimited output
- **`outlook/groups.py`**: AsyncGroup class for handling async Click commands
- **`outlook/clients/__init__.py`**: OutlookClient class - main interface for Graph API operations
- **`outlook/clients/auth.py`**: GraphAuthClient for Azure AD authentication with descriptor pattern
- **`outlook/clients/folders.py`**: Folders class - dictionary-like collection of mail folders
- **`outlook/clients/users.py`**: User dataclass for caching user profile information
- **`outlook/clients/settings.py`**: Configuration file path definitions
- **`.auth.json`**: Azure AD app configuration (user-created)
- **`.auth_record.json`**: Cached authentication record (auto-generated)

### Key Design Patterns

- **Background Loading**: OutlookClient starts background tasks to load user and folder data in parallel for improved performance
- **Descriptor Pattern**: AuthenticationRecordDescriptor manages persistent authentication records with automatic serialization/deserialization
- **Async/Await**: All Graph API operations use async/await for efficient I/O
- **AsyncGroup**: Custom Click group class that provides `async_command()` decorator to seamlessly integrate async functions with Click CLI
- **Pipe-Delimited Output**: All commands output structured data in pipe-delimited format for easy parsing and integration
