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
)
from .groups import AsyncGroup


@click.group(cls=AsyncGroup)
def cli() -> None:
    """Microsoft Graph API CLI for Outlook/Office 365"""
    pass


@cli.async_command()
async def user() -> None:
    """Display current user information"""
    try:
        if user_info := await get_user():
            click.echo(f"Hello, {user_info.display_name}")
            # For Work/school accounts, email is in mail property
            # Personal accounts, email is in userPrincipalName
            click.echo(f"Email: {user_info.mail or user_info.user_principal_name}")
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
        click.echo("Successfully authenticated and saved credentials")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


@cli.async_command()
async def folders() -> None:
    """List all mail folders"""
    try:
        folder_response = await get_folders()
        if folder_response and folder_response.value:
            for folder in folder_response.value:
                click.echo(f"{folder.display_name}: {folder.id}")
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
                click.echo(f"ID: {message.id}")
                click.echo(f"Subject: {message.subject}")
                if message.from_ and message.from_.email_address:
                    click.echo(f"From: {message.from_.email_address.name or 'NONE'}")
                else:
                    click.echo("From: NONE")
                click.echo(f"Status: {'Read' if message.is_read else 'Unread'}")
                click.echo(f"Received: {message.received_date_time}")
                click.echo()

            more_available = message_page.odata_next_link is not None
            click.echo(f"More messages available? {more_available}")
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
            click.echo(f"Subject: {message.subject}")
            if message.from_ and message.from_.email_address:
                click.echo(
                    f"From: {message.from_.email_address.name or 'NONE'} <{message.from_.email_address.address}>"
                )
            else:
                click.echo("From: NONE")
            click.echo(f"Received: {message.received_date_time}")
            click.echo(f"Status: {'Read' if message.is_read else 'Unread'}")
            click.echo("\n--- Body ---")
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
        click.echo(f"Message {message_id} moved successfully")
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
        click.echo(f"Message {message_id} moved to Deleted Items")
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
        click.echo(f"Message {message_id} forwarded to {', '.join(recipients)}")
    except ODataError as e:
        click.echo(
            f"Error: {e.error.code if e.error else 'Unknown'} - {e.error.message if e.error else ''}",
            err=True,
        )


if __name__ == "__main__":
    cli()
