# Python Outlook Graph API Tutorial

## Project Overview

This is a Python application that interacts with Microsoft Graph API to access Outlook/Office 365 data. It provides a command-line interface for authenticating with Azure AD and performing operations on a user's mailbox including reading, moving, deleting, and forwarding emails.

## Architecture

### Main Components

**`outlook/__main__.py`** - Main entry point
- CLI using Click with AsyncGroup for async command support
- Implements async/await pattern throughout
- Handles OData errors from Graph API
- Commands: login, user, folders, list, read, move, delete, forward

**`outlook/groups.py`** - AsyncGroup class
- Custom Click Group that provides `async_command()` decorator
- Automatically wraps async functions with `asyncio.run()`
- Seamlessly integrates async/await with Click's command system

**`outlook/actions/__init__.py`** - Action functions
- Pure async functions for Graph API operations
- Decorated with `@Client.decorator` for automatic client management
- Functions:
  - `authenticate()` - Authenticate and save credentials
  - `get_user()` - Fetches user profile
  - `get_folders()` - Lists all mail folders
  - `get_messages()` - Retrieves messages from a folder with filtering/sorting
  - `get_message()` - Fetches a single message with full body
  - `message_move()` - Moves message to different folder
  - `message_delete()` - Soft deletes message (moves to Deleted Items)
  - `message_forward()` - Forwards message to recipients
  - `folder_create()` - Creates new mail folder

**`outlook/actions/clients.py`** - Client management
- `Client` class with context manager pattern
- `InteractiveBrowserCredential` with token caching and Windows Account Manager (WAM) support
- `Client.decorator` wraps action functions to provide Graph client automatically
- Loads configuration from `.auth.json`
- Manages authentication records in `.auth_record.json`

### Features

- Silent authentication with token caching and Windows Account Manager
- Browser-based OAuth authentication (InteractiveBrowserCredential)
- View user profile information
- List and manage mail folders
- List messages with filtering, sorting, and pagination
- Read full message content including HTML/text body
- Move messages between folders
- Delete messages (soft delete)
- Forward messages to multiple recipients with optional comments
- Async/await architecture throughout
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

# User info
python -m outlook user

# Mail folders
python -m outlook folders

# List messages
python -m outlook list [folder_id] [--limit N]

# Read message
python -m outlook read <message_id>

# Move message
python -m outlook move <message_id> <destination_folder_id>

# Delete message
python -m outlook delete <message_id>

# Forward message
python -m outlook forward <message_id> <recipients...> [--comment TEXT]
```

## Key Design Patterns

### Client Decorator Pattern

The `Client.decorator` provides automatic Graph client management:

```python
@Client.decorator
async def get_user(client: GraphServiceClient) -> User | None:
    return await client.me.get(...)
```

The decorator:
1. Creates a `Client` instance with authentication
2. Provides the `GraphServiceClient` as the first parameter
3. Handles context management and cleanup
4. Enables reusable action functions without boilerplate

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

### Token Caching with WAM

The `InteractiveBrowserCredential` uses:
- **Token cache persistence** - Avoids repeated authentication
- **Windows Account Manager (WAM)** - Native Windows integration for silent auth
- **Authentication records** - Stored in `.auth_record.json` for session persistence

## Dependencies

- `azure-identity` - Azure authentication with WAM support
- `msgraph-sdk` - Microsoft Graph SDK for Python
- `click` - Command-line interface framework
- Python 3.13+ (for async/await and modern type hints)
