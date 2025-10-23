from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Final

import click

from .groups import AsyncGroup
from .utils import get_emails_str, get_from_str, sanitize_for_output

from .clients import OutlookClient
from .purge import purge_worker

TORONTO: Final[ZoneInfo] = ZoneInfo("America/Toronto")


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
        now_toronto = datetime.now(TORONTO)
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


@cli.command()
@click.argument("folder_id", default="inbox")
@click.option("--dry-run", is_flag=True, help="Show what would be deleted without actually deleting")
@click.option("--batch-size", default=50, help="Number of emails to process in each batch")
@click.option("--before-date", default="2024-01-01", help="Delete emails received before this date (YYYY-MM-DD)")
def purge(folder_id: str, dry_run: bool, batch_size: int, before_date: str) -> None:
    """Delete old emails from a folder in batches"""

    # Parse and validate the date
    try:
        cutoff_date = datetime.fromisoformat(before_date)
        if cutoff_date.tzinfo is None:
            cutoff_date = cutoff_date.replace(tzinfo=TORONTO)
    except ValueError:
        click.echo(f"ERROR: Invalid date format '{before_date}'. Use YYYY-MM-DD")
        return

    # Print header
    mode = "DRY-RUN: " if dry_run else ""
    click.echo(f"{mode}Purging emails from folder '{folder_id}' before {before_date}")
    click.echo(f"Batch size: {batch_size}")
    click.echo("Type 'QUIT' (case-insensitive) and press Enter to finish current batch and exit")
    click.echo("-" * 60)

    # Call the purge worker (handles threading internally)
    purge_worker(folder_id, batch_size=batch_size, before_date=cutoff_date, dry_run=dry_run)


if __name__ == "__main__":
    cli()
