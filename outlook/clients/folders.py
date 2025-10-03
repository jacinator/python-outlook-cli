from __future__ import annotations

from collections.abc import Generator
from typing import ClassVar

from msgraph.generated.models.mail_folder import MailFolder
from msgraph.generated.models.mail_folder_collection_response import MailFolderCollectionResponse
from msgraph.graph_service_client import GraphServiceClient


class Folders(dict[str, MailFolder]):
    """Store folder objects with their IDs as keys for quick access.

    Store full MailFolder objects so that folder metadata (display name, child
    count, total items, etc.) can be accessed without additional API calls.
    """

    EXCLUDED_FOLDERS: ClassVar[frozenset[str]] = frozenset(
        {"Conversation History", "Drafts", "Outbox", "RSS Subscriptions", "Sent Items"}
    )

    @staticmethod
    def _get_nested_folders(folders: list[MailFolder] | None) -> Generator[tuple[str, MailFolder]]:
        if not folders:
            return

        for folder in folders:
            if (
                not folder.id
                or not folder.display_name
                or folder.display_name in Folders.EXCLUDED_FOLDERS
            ):
                continue

            yield from Folders._get_nested_folders(folder.child_folders)
            yield (folder.id, folder)

    @classmethod
    async def create(cls, client: GraphServiceClient) -> Folders:
        """Load all folders from the mailbox and return a Folders instance.

        Load all the folders through Microsoft Graph and map their IDs to
        the full MailFolder objects. This allows the program to access folder
        metadata without calling the API again.
        """
        folders: MailFolderCollectionResponse | None = await client.me.mail_folders.get()
        return cls(cls._get_nested_folders(getattr(folders, "value", None)))
