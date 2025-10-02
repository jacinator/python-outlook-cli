import click
from msgraph.generated.models.o_data_errors.o_data_error import ODataError

from .actions import (
    authenticate,
    get_folders,
    get_message,
    get_messages,
    get_user,
    message_delete,
    message_forward,
    message_move,
    message_move_all,
)
from .groups import AsyncGroup
from .utils import get_emails_str, get_from_str


@click.group(cls=AsyncGroup)
def cli() -> None:
    """Microsoft Graph API CLI for Outlook/Office 365"""
    pass


@cli.async_command()
async def user() -> None:
    """Display current user information"""
    try:
        if user_info := await get_user():
            # For Work/school accounts, email is in mail property
            # Personal accounts, email is in userPrincipalName
            email = user_info.mail or user_info.user_principal_name or "NONE"
            display_name = user_info.display_name or "NONE"
            click.echo(f"{display_name}|{email}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.command()
def login() -> None:
    """Authenticate and save credentials"""
    try:
        authenticate()
        click.echo("OK|authenticated")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("folder_id", required=False)
async def folders(folder_id: str | None) -> None:
    """List mail folders (top-level or child folders of specified parent)"""
    try:
        folder_response = await get_folders(folder_id)
        if folder_response and folder_response.value:
            for folder in folder_response.value:
                parts = [
                    folder.display_name or "NONE",
                    folder.id or "NONE",
                    f"parent={folder.parent_folder_id or 'NONE'}",
                    f"children={folder.child_folder_count if folder.child_folder_count is not None else 0}",
                    f"total={folder.total_item_count if folder.total_item_count is not None else 0}",
                    f"unread={folder.unread_item_count if folder.unread_item_count is not None else 0}",
                    f"hidden={str(folder.is_hidden).lower() if folder.is_hidden is not None else 'false'}",
                ]
                click.echo("|".join(parts))
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("folder_id", default="inbox")
@click.option("--limit", "-l", default=25, help="Number of messages to retrieve")
async def list(folder_id: str, limit: int) -> None:
    """List messages in a folder"""
    try:
        message_page = await get_messages(folder_id, top=limit)
        if message_page and message_page.value:
            for message in message_page.value:
                from_str: str = get_from_str(message.from_)
                to_str: str = get_emails_str(message.to_recipients)
                cc_str: str = get_emails_str(message.cc_recipients)

                parts = [
                    message.id or "NONE",
                    message.subject or "NONE",
                    from_str,
                    f"to={to_str}",
                    f"cc={cc_str}",
                    "read" if message.is_read else "unread",
                    str(message.received_date_time) if message.received_date_time else "NONE",
                    f"sent={message.sent_date_time}" if message.sent_date_time else "sent=NONE",
                    f"attachments={str(message.has_attachments).lower() if message.has_attachments is not None else 'false'}",
                    f"importance={message.importance.value if message.importance else 'normal'}",
                    f"conversation={message.conversation_id or 'NONE'}",
                    f"folder={message.parent_folder_id or 'NONE'}",
                    f"weblink={message.web_link or 'NONE'}",
                ]
                click.echo("|".join(parts))

            more_available = message_page.odata_next_link is not None
            click.echo(f"--- more={str(more_available).lower()} ---")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("message_id")
async def read(message_id: str) -> None:
    """Display full message including body"""
    try:
        if message := await get_message(message_id):
            from_str: str = get_from_str(message.from_)
            to_str: str = get_emails_str(message.to_recipients)
            cc_str: str = get_emails_str(message.cc_recipients)

            # Build header parts
            header_parts = [
                f"id={message.id or 'NONE'}",
                f"subject={message.subject or 'NONE'}",
                f"from={from_str}",
                f"to={to_str}",
                f"cc={cc_str}",
                f"received={message.received_date_time}" if message.received_date_time else "received=NONE",
                f"sent={message.sent_date_time}" if message.sent_date_time else "sent=NONE",
                f"status={'read' if message.is_read else 'unread'}",
                f"attachments={str(message.has_attachments).lower() if message.has_attachments is not None else 'false'}",
                f"importance={message.importance.value if message.importance else 'normal'}",
                f"conversation={message.conversation_id or 'NONE'}",
                f"folder={message.parent_folder_id or 'NONE'}",
                f"weblink={message.web_link or 'NONE'}",
            ]
            click.echo("|".join(header_parts))

            # Body
            body_type = message.body.content_type.value if message.body and message.body.content_type else "text"
            click.echo(f"\n--- Body ({body_type}) ---")
            if message.body and message.body.content:
                click.echo(message.body.content)
            else:
                click.echo("(No body content)")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("message_id")
@click.argument("destination_folder_id")
async def move(message_id: str, destination_folder_id: str) -> None:
    """Move a message to a different folder"""
    try:
        await message_move(message_id, destination_folder_id)
        click.echo(f"OK|moved|{message_id}|to|{destination_folder_id}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("source_folder_id")
@click.argument("destination_folder_id")
async def moveall(source_folder_id: str, destination_folder_id: str) -> None:
    """Move all messages from source folder to destination folder"""
    try:
        count = await message_move_all(source_folder_id, destination_folder_id)
        click.echo(f"OK|moved|count={count}|from|{source_folder_id}|to|{destination_folder_id}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("message_id")
async def delete(message_id: str) -> None:
    """Delete a message (moves to Deleted Items)"""
    try:
        await message_delete(message_id)
        click.echo(f"OK|deleted|{message_id}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
@click.argument("message_id")
@click.argument("recipients", nargs=-1, required=True)
@click.option("--comment", "-c", help="Comment to add when forwarding")
async def forward(
    message_id: str, recipients: tuple[str, ...], comment: str | None
) -> None:
    """Forward a message to one or more recipients"""
    try:
        await message_forward(message_id, list(recipients), comment)
        recipients_str = ",".join(recipients)
        click.echo(f"OK|forwarded|{message_id}|to|{recipients_str}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


if __name__ == "__main__":
    cli()
