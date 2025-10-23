# Python Outlook Graph API Tutorial

## Project Overview

This is a Python application that interacts with Microsoft Graph API to access Outlook/Office 365 data. It provides a command-line interface for authenticating with Azure AD and performing operations on a user's mailbox including reading, moving, and deleting emails. The CLI uses pipe-delimited output format for easy integration with scripts and AI tools.

## Architecture

### Main Components

**`outlook/__main__.py`** - Main entry point
- CLI using Click with AsyncGroup for async command support
- Implements async/await pattern throughout
- Handles OData errors from Graph API
- Commands: login, user, folders, list, read, move, delete, purge
- Output format: Pipe-delimited for AI/script-friendly parsing

**`outlook/groups.py`** - AsyncGroup class
- Custom Click Group that provides `async_command()` decorator
- Automatically wraps async functions with `asyncio.run()`
- Seamlessly integrates async/await with Click's command system

**`outlook/clients/__init__.py`** - OutlookClient class
- Main interface for Graph API operations
- Manages authentication via `GraphAuthClient`
- Background loading of user and folder data for performance
- Methods:
  - `authenticate()` - Perform interactive browser authentication
  - `user()` - Get cached user profile
  - `folders()` - Get cached folder structure
  - `get_messages()` - Retrieve messages with filtering/sorting
  - `get_message()` - Fetch single message with full body
  - `move_messages()` - Move one or more messages to different folder (parallel execution)
  - `delete_messages()` - Soft delete one or more messages (moves to Deleted Items, parallel execution)

**`outlook/clients/auth.py`** - Authentication management
- `GraphAuthClient` - Handles Azure AD authentication
- `Config` - Stores client ID, tenant ID, and scopes
- `AuthenticationRecordDescriptor` - Descriptor for persisting auth records
- `InteractiveBrowserCredential` with token caching and Windows Account Manager (WAM) support
- Loads configuration from `.auth.json`
- Manages authentication records in `.auth_record.json`

**`outlook/clients/folders.py`** - Folder management
- `Folders` - Dictionary-like collection of mail folders
- Stores full `MailFolder` objects for metadata access
- Recursively loads nested folders
- Excludes system folders (Drafts, Sent Items, etc.)

**`outlook/clients/users.py`** - User profile management
- `User` - Dataclass for storing user profile information
- Caches display name and email address
- Handles both work/school and personal account formats

**`outlook/clients/settings.py`** - Configuration paths
- Defines paths for `.auth.json` and `.auth_record.json`
- Provides root directory reference

**`outlook/purge.py`** - Bulk email purge functionality
- `purge_worker()` - Main function for batch email deletion
- Runs in background thread with foreground user interaction
- Uses single event loop per worker thread for efficient async operations
- Progress reporting with `PROGRESS|counter|total|batch` format
- Final summary with `RESULT|action|total|folder_id|before_date` format
- Interactive cancellation via "QUIT" command
- Supports dry-run mode for testing

### Features

- Silent authentication with token caching and Windows Account Manager (WAM)
- Browser-based OAuth authentication (InteractiveBrowserCredential)
- View user profile information
- List and manage mail folders
- List messages with filtering, sorting, and pagination
- Read full message content including HTML/text body
- Move messages between folders
- Delete messages (soft delete to Deleted Items)
- Bulk purge old emails with interactive cancellation
- Async/await architecture throughout
- Background data loading for improved performance
- Pipe-delimited output format for AI/script integration
- Error handling for OData errors

### Configuration Required

Create a `.auth.json` file with:
```json
{
  "clientId": "YOUR_AZURE_AD_APP_ID",
  "tenantId": "YOUR_AZURE_AD_TENANT_ID",
  "graphUserScopes": "User.Read Mail.ReadWrite Mail.Send"
}
```

## Running the Program

### From Virtual Environment

Activate the virtual environment and run:

```bash
# Activate venv (Linux/WSL)
source venv/bin/activate

# Or on Windows
venv\Scripts\activate

# Run commands
python -m outlook <command>
```

### From Outside Virtual Environment

```bash
# Linux/WSL
./venv/bin/python -m outlook <command>

# Windows
.\venv\Scripts\python -m outlook <command>
```

### Available Commands

```bash
# Authentication
python -m outlook login

# User info (pipe-delimited: name|email)
python -m outlook user

# Mail folders (pipe-delimited with metadata)
python -m outlook folders

# List messages (pipe-delimited with full metadata)
python -m outlook list [folder_id] [--top N] [--oldest-first]
python -m outlook list inbox --top 100  # List 100 messages from inbox
python -m outlook list inbox --oldest-first  # List oldest messages first

# Read message (pipe-delimited header + body)
python -m outlook read <message_id>

# Move message(s)
python -m outlook move <message_id> <destination_folder_id>
python -m outlook move <message_id1> <message_id2> <message_id3> <destination_folder_id>  # Batch move

# Delete message(s) (soft delete to Deleted Items)
python -m outlook delete <message_id>
python -m outlook delete <message_id1> <message_id2> <message_id3>  # Batch delete

# Purge old emails (bulk deletion with interactive control)
python -m outlook purge [folder_id] [--dry-run] [--batch-size N] [--before-date YYYY-MM-DD]
python -m outlook purge inbox --dry-run --batch-size 50  # Test run, 50 emails per batch
python -m outlook purge inbox --before-date 2023-01-01  # Delete emails before Jan 1, 2023
python -m outlook purge --dry-run  # Test run on inbox with defaults (before 2024-01-01)
# Type 'QUIT' (case-insensitive) and press Enter to stop gracefully
```

## Key Design Patterns

### OutlookClient with Background Loading

The `OutlookClient` class provides a high-level interface with optimized data loading:

```python
# Client initialization starts background tasks
manager = OutlookClient()

# Data is loaded in parallel, awaited only when accessed
user = await manager.user()  # Returns cached user data
folders = await manager.folders()  # Returns cached folder data
messages, more = await manager.get_messages("inbox", top=50)
```

Key features:
1. Initializes `GraphAuthClient` for authentication
2. Starts background tasks to load user and folder data in parallel
3. Data is awaited only when accessed, improving performance
4. Provides typed methods for all message operations

### AsyncGroup for Click Commands

Custom Click group that bridges async/await with Click's command system:

```python
@cli.async_command()
async def user() -> None:
    user_info = await get_user()
    click.echo(f"Hello, {user_info.display_name}")
```

The `async_command()` decorator:
1. Wraps the async function with `asyncio.run()`
2. Registers it as a Click command
3. Preserves function metadata
4. Allows natural async/await syntax in CLI commands

### Authentication Record Descriptor

The `AuthenticationRecordDescriptor` manages persistent authentication:

```python
class GraphAuthClient:
    auth: ClassVar[AuthenticationRecordDescriptor] = AuthenticationRecordDescriptor()
```

The descriptor:
1. Automatically loads auth records from `.auth_record.json` on access
2. Serializes and saves auth records to disk on assignment
3. Enables deletion of auth records to force re-authentication
4. Works seamlessly with `InteractiveBrowserCredential`

### Token Caching with WAM

The `InteractiveBrowserCredential` uses:
- **Token cache persistence** - Avoids repeated authentication
- **Windows Account Manager (WAM)** - Native Windows integration for silent auth
- **Authentication records** - Stored in `.auth_record.json` for session persistence via descriptor

### Pipe-Delimited Output Format

All commands output pipe-delimited data for easy parsing by scripts and AI:

```bash
# User command output
John Doe|john.doe@example.com

# Folders command output
Inbox|AAMkAD...|parent=NONE|children=5|total=342|unread=12|hidden=false

# List command output
AAMkAD...|Meeting Tomorrow|Jane Smith <jane@example.com>|to=...|cc=...|read|2025-10-01 14:26:09+00:00|...
```

This format enables:
- Easy parsing with `split('|')`
- Structured data extraction
- AI-friendly processing
- Script automation

## Dependencies

- `azure-identity` - Azure authentication with WAM support
- `msgraph-sdk` - Microsoft Graph SDK for Python
- `click` - Command-line interface framework
- Python 3.13+ (for async/await and modern type hints)
