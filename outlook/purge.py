import asyncio
from datetime import datetime
from threading import Event, Thread

import click
from msgraph.generated.models.message import Message

from .clients import OutlookClient
from .utils import sanitize_for_output


def purge_worker(
    folder_id: str, *, batch_size: int, before_date: datetime, dry_run: bool
) -> None:
    quit_event: Event = Event()

    async def _purge_iteration(iteration: int, total: int) -> int:
        manager: OutlookClient = OutlookClient()

        messages: list[Message]
        messages, _ = await manager.get_messages(
            folder_id,
            top=batch_size,
            orderby=("receivedDateTime ASC",),
            filter=f"receivedDateTime lt {before_date.isoformat()}",
        )

        if not messages:
            click.echo("No more old emails found. Purge complete!")
            quit_event.set()
            return 0

        message_queue: list[tuple[str, str | None, datetime | None]] = [
            (x.id, x.subject, x.received_date_time) for x in messages if x.id
        ]

        action: str = "DRY-RUN" if dry_run else "DELETING"
        for m_id, m_subject, m_received in message_queue:
            click.echo(
                "{}|{}|{}|{}".format(
                    action,
                    m_id,
                    sanitize_for_output(m_subject or "NONE"),
                    m_received or "NONE",
                )
            )

        batch: int = len(message_queue)
        total += batch
        click.echo(f"PROGRESS|{total=}|{iteration=}|{batch=}")

        if not dry_run:
            await manager.delete_messages(tuple(x[0] for x in message_queue))

        return batch

    def _purge_worker() -> None:
        counter: int = 1
        total: int = 0

        # Create ONE event loop for the entire worker thread
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)

        try:
            while not quit_event.is_set():
                # Reuse the same loop for each batch
                total += loop.run_until_complete(_purge_iteration(counter, total))
                counter += 1
        finally:
            # Clean up the loop when done
            loop.close()

        action: str = "DRY-RUN" if dry_run else "DELETED"
        click.echo(f"RESULT|{action}|{total=}|{folder_id=}|before={before_date.isoformat()}")

    thread: Thread = Thread(target=_purge_worker, daemon=False)
    thread.start()

    while thread.is_alive():
        if click.prompt("", prompt_suffix="", value_proc=str.upper) == "QUIT":
            click.echo("User requested quit. Finishing current batch...")
            quit_event.set()
            break
