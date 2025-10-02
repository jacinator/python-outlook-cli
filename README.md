# Python Outlook Graph API CLI

A command-line interface for managing Microsoft Outlook emails using the Microsoft Graph API.

## Features

- Silent authentication with Windows Account Manager (WAM) and token caching
- Read, list, move, delete, and forward emails
- Manage mail folders
- Full email body access
- Built with async/await for efficient API calls

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

### User Information

#### `user`
Display current user information.

```bash
python -m outlook user
```

Output:
```
Hello, John Doe
Email: john.doe@example.com
```

### Managing Folders

#### `folders`
List all mail folders with their IDs.

```bash
python -m outlook folders
```

Output:
```
Inbox: AAMkAD...
Sent Items: AAMkAD...
Drafts: AAMkAD...
Deleted Items: AAMkAD...
```

### Listing Messages

#### `list [FOLDER_ID] [--limit N]`
List messages from a specific folder.

```bash
# List inbox (default)
python -m outlook list

# List specific folder
python -m outlook list AAMkAD...

# List 50 messages
python -m outlook list inbox --limit 50
python -m outlook list --limit 50  # Short form
```

Output includes:
- Message ID
- Subject
- From
- Read/Unread status
- Received date/time

### Reading Messages

#### `read MESSAGE_ID`
Display full message including body content.

```bash
python -m outlook read AAMkAD...
```

Output:
```
Subject: Meeting Tomorrow
From: Jane Smith <jane@example.com>
Received: 2025-10-01 14:26:09+00:00
Status: Read

--- Body ---
Hi team,
Let's meet tomorrow at 2pm...
```

### Moving Messages

#### `move MESSAGE_ID DESTINATION_FOLDER_ID`
Move a message to a different folder.

```bash
python -m outlook move AAMkAD... AAMkAD...
```

### Deleting Messages

#### `delete MESSAGE_ID`
Move a message to Deleted Items folder (soft delete).

```bash
python -m outlook delete AAMkAD...
```

**Note**: This performs a soft delete (moves to Deleted Items). Permanent deletion is not available via the CLI.

### Forwarding Messages

#### `forward MESSAGE_ID RECIPIENT [RECIPIENT...] [--comment TEXT]`
Forward a message to one or more recipients.

```bash
# Forward to one recipient
python -m outlook forward AAMkAD... john@example.com

# Forward to multiple recipients
python -m outlook forward AAMkAD... john@example.com jane@example.com

# Forward with comment
python -m outlook forward AAMkAD... john@example.com --comment "FYI"
python -m outlook forward AAMkAD... john@example.com -c "Please review"
```

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
venv\Scripts\python.exe -m src <command>
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

- **`outlook/__main__.py`**: CLI commands using Click with async support
- **`outlook/groups.py`**: AsyncGroup class for handling async Click commands
- **`outlook/actions/__init__.py`**: Action functions for Microsoft Graph API operations
- **`outlook/actions/clients.py`**: Client class with decorator pattern for authentication and Graph API client management
- **`.auth.json`**: Azure AD app configuration (user-created)
- **`.auth_record.json`**: Cached authentication record (auto-generated)

### Key Design Patterns

- **Decorator Pattern**: The `Client.decorator` wraps action functions to automatically handle Graph client initialization and context management
- **Async/Await**: All Graph API operations use async/await for efficient I/O
- **AsyncGroup**: Custom Click group class that provides `async_command()` decorator to seamlessly integrate async functions with Click CLI
