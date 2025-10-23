import click
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from .groups import AsyncGroup
from .utils import get_emails_str, get_from_str, sanitize_for_output

from .clients import OutlookClient


@click.group(cls=AsyncGroup)
def cli() -> None:
    """Microsoft Graph API CLI for Outlook/Office 365"""
    pass


@cli.command()
def login() -> None:
    """Authenticate and save credentials"""
    manager = OutlookClient()
    manager.authenticate()
    click.echo("OK|authenticated")


@cli.async_command()
async def user() -> None:
    """Display current user information from cached data"""
    manager = OutlookClient()
    user = await manager.user()

    click.echo("{}|{}".format(user.name or "NONE", user.addr or "NONE"))


@cli.async_command()
async def folders() -> None:
    """List all mail folders from cached data"""
    manager = OutlookClient()
    folders = await manager.folders()

    click.echo(
        "\n".join(
            "|".join(
                (
                    folder.display_name or "NONE",
                    folder.id or "NONE",
                    f"parent={folder.parent_folder_id or 'NONE'}",
                    f"children={folder.child_folder_count or 0}",
                    f"total={folder.total_item_count or 0}",
                    f"unread={folder.unread_item_count or 0}",
                    f"hidden={str(folder.is_hidden).lower() or 'false'}",
                )
            )
            for folder in folders.values()
        )
    )


@cli.async_command()
@click.argument("folder_id", default="inbox")
@click.option("--top", "-t", default=100, help="Number of messages to retrieve")
@click.option("--oldest-first", is_flag=True, help="Sort messages oldest first instead of newest first")
@click.option("--today", is_flag=True, help="Filter messages received today (America/Toronto timezone)")
@click.option("--yesterday", is_flag=True, help="Filter messages received yesterday (America/Toronto timezone)")
async def list(folder_id: str, top: int, oldest_first: bool, today: bool, yesterday: bool) -> None:
    """List messages in a folder"""

    # Validate mutually exclusive flags
    if today and yesterday:
        raise click.UsageError("--today and --yesterday cannot be used together")

    # Build date filter if needed
    filter_expr: str | None = None
    if today or yesterday:
        toronto_tz = ZoneInfo("America/Toronto")
        now_toronto = datetime.now(toronto_tz)
        today_start = now_toronto.replace(hour=0, minute=0, second=0, microsecond=0)

        if today:
            start_time = today_start
            end_time = today_start + timedelta(days=1)
        else:  # yesterday
            start_time = today_start - timedelta(days=1)
            end_time = today_start

        # Format as ISO 8601 for OData
        filter_expr = f"receivedDateTime ge {start_time.isoformat()} and receivedDateTime lt {end_time.isoformat()}"

    orderby = "ASC" if oldest_first else "DESC"
    manager = OutlookClient()
    messages, more_available = await manager.get_messages(folder_id, top=top, orderby=(f"receivedDateTime {orderby}",), filter=filter_expr)
    for message in messages:
        from_str: str = get_from_str(message.from_)
        to_str: str = get_emails_str(message.to_recipients)
        cc_str: str = get_emails_str(message.cc_recipients)

        parts = [
            message.id or "NONE",
            sanitize_for_output(message.subject or "NONE"),
            sanitize_for_output(from_str),
            f"to={sanitize_for_output(to_str)}",
            f"cc={sanitize_for_output(cc_str)}",
            "read" if message.is_read else "unread",
            str(message.received_date_time) if message.received_date_time else "NONE",
            f"sent={message.sent_date_time}" if message.sent_date_time else "sent=NONE",
            f"attachments={str(message.has_attachments).lower() if message.has_attachments is not None else 'false'}",
            f"importance={message.importance.value if message.importance else 'normal'}",
            f"conversation={message.conversation_id or 'NONE'}",
            f"folder={message.parent_folder_id or 'NONE'}",
            f"weblink={message.web_link or 'NONE'}",
        ]
        click.echo(sanitize_for_output("|".join(parts)))
    click.echo(f"--- more={more_available} ---".lower())


@cli.async_command()
@click.argument("message_id")
async def read(message_id: str) -> None:
    """Display full message including body"""
    manager = OutlookClient()
    if message := await manager.get_message(message_id):
        from_str: str = get_from_str(message.from_)
        to_str: str = get_emails_str(message.to_recipients)
        cc_str: str = get_emails_str(message.cc_recipients)

        # Build header parts
        header_parts = [
            f"id={message.id or 'NONE'}",
            f"subject={sanitize_for_output(message.subject or 'NONE')}",
            f"from={sanitize_for_output(from_str)}",
            f"to={sanitize_for_output(to_str)}",
            f"cc={sanitize_for_output(cc_str)}",
            f"received={message.received_date_time}" if message.received_date_time else "received=NONE",
            f"sent={message.sent_date_time}" if message.sent_date_time else "sent=NONE",
            f"status={'read' if message.is_read else 'unread'}",
            f"attachments={str(message.has_attachments).lower() if message.has_attachments is not None else 'false'}",
            f"importance={message.importance.value if message.importance else 'normal'}",
            f"conversation={message.conversation_id or 'NONE'}",
            f"folder={message.parent_folder_id or 'NONE'}",
            f"weblink={message.web_link or 'NONE'}",
        ]
        click.echo(sanitize_for_output("|".join(header_parts)))

        # Body
        body_type = message.body.content_type.value if message.body and message.body.content_type else "text"
        click.echo(f"\n--- Body ({body_type}) ---")
        if message.body and message.body.content:
            click.echo(sanitize_for_output(message.body.content))
        else:
            click.echo("(No body content)")


@cli.async_command()
@click.argument("message_ids", nargs=-1, required=True)
@click.argument("folder_id")
async def move(message_ids: tuple[str, ...], folder_id: str) -> None:
    """Move one or more messages to a different folder"""
    manager = OutlookClient()
    await manager.move_messages(folder_id, message_ids)
    click.echo("\n".join(f"OK|moved|{message_id}|to|{folder_id}" for message_id in message_ids))


@cli.async_command()
@click.argument("message_ids", nargs=-1, required=True)
async def delete(message_ids: tuple[str, ...]) -> None:
    """Delete one or more messages (moves to Deleted Items)"""
    manager = OutlookClient()
    await manager.delete_messages(message_ids)
    click.echo("\n".join(f"OK|deleted|{message_id}" for message_id in message_ids))


if __name__ == "__main__":
    cli()
