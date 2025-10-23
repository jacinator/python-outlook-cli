from asyncio import Task, create_task, gather, Semaphore
from dataclasses import dataclass
from typing import Final

from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.mail_folders.item.mail_folder_item_request_builder import MailFolderItemRequestBuilder
from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody

from .auth import GraphAuthClient
from .folders import Folders
from .users import User

DEFAULT_ORDERBY: Final[tuple[str, ...]] = ("receivedDateTime DESC",)
DEFAULT_SELECT: Final[tuple[str, ...]] = ("id", "subject", "from", "toRecipients", "ccRecipients", "isRead", "receivedDateTime", "sentDateTime", "hasAttachments", "importance", "conversationId", "parentFolderId", "webLink", "bodyPreview")


@dataclass(slots=True)
class OutlookClient:
    """Manager for Microsoft Graph API operations with cached data.

    Provides a higher-level interface for interacting with Microsoft Graph API
    by maintaining authenticated client and pre-loaded folder information to
    minimize repeated API calls.

    User and folder data load in the background when Manager is created, and
    are awaited only when accessed via user() or folders() methods.
    """

    auth: GraphAuthClient
    client: GraphServiceClient

    _folders_task: Task[Folders]
    _user_task: Task[User]

    def __init__(self) -> None:
        self.auth = GraphAuthClient()
        self.client = self.auth.client

        self._folders_task = create_task(Folders.create(self.client))
        self._user_task = create_task(User.create(self.client))

    def authenticate(self) -> None:
        """Perform interactive browser authentication and save credentials.

        Opens a browser window for the user to authenticate with Microsoft
        and saves the authentication record for future silent authentication.
        """
        self.auth.auth = self.auth.credentials.authenticate(scopes=self.auth.config.scopes)

    # =========================================================================
    # Cached Data Access Methods
    # =========================================================================
    # These methods provide access to user and folder data that is loaded
    # asynchronously in the background when OutlookClient is instantiated.
    # This approach improves performance by parallelizing data fetching rather
    # than blocking on each request sequentially.

    async def folders(self) -> Folders:
        """Get all mail folders for the authenticated user.

        Returns cached folder data loaded in the background during initialization.
        If the background task hasn't completed, this will wait for it.

        Returns:
            Folders: Dictionary-like collection of mail folders keyed by folder ID.
        """
        return await self._folders_task

    async def user(self) -> User:
        """Get profile information for the authenticated user.

        Returns cached user data loaded in the background during initialization.
        If the background task hasn't completed, this will wait for it.

        Returns:
            User: User profile containing name, email, and other account details.
        """
        return await self._user_task

    # =========================================================================
    # Internal Helper Methods
    # =========================================================================
    # These private methods provide typed request builders for Graph API
    # resources, reducing code duplication across public methods.

    def _get_folder(self, folder_id: str) -> MailFolderItemRequestBuilder:
        """Get a request builder for a specific mail folder.

        Args:
            folder_id: The ID of the mail folder to access.

        Returns:
            MailFolderItemRequestBuilder: Builder for performing operations on the folder.
        """
        return self.client.me.mail_folders.by_mail_folder_id(folder_id)

    def _get_message(self, message_id: str) -> MessageItemRequestBuilder:
        """Get a request builder for a specific message.

        Args:
            message_id: The ID of the message to access.

        Returns:
            MessageItemRequestBuilder: Builder for performing operations on the message.
        """
        return self.client.me.messages.by_message_id(message_id)

    # =========================================================================
    # Message Operations
    # =========================================================================
    # These methods provide high-level operations for working with email
    # messages, including retrieval, modification, and deletion. All methods
    # operate asynchronously and interact with the Microsoft Graph API.

    async def delete_messages(self, message_ids: tuple[str, ...]) -> None:
        """Delete multiple messages in parallel (moves to Deleted Items folder).

        Performs a soft delete by moving the messages to the Deleted Items folder
        rather than permanently removing them. Operations are executed in parallel
        with a maximum of 4 concurrent requests to comply with API guidelines.

        Args:
            message_ids: Tuple of message IDs to delete.
        """
        semaphore = Semaphore(4)

        async def delete_with_semaphore(message_id: str) -> None:
            async with semaphore:
                await self._get_message(message_id).delete()

        await gather(*(delete_with_semaphore(message_id) for message_id in message_ids))

    async def get_message(self, message_id: str) -> Message | None:
        """Retrieve a single message by ID with full content.

        Fetches complete message details including headers, body content,
        recipients, and metadata.

        Args:
            message_id: The ID of the message to retrieve.

        Returns:
            Message | None: The message object if found, None otherwise.
        """
        message: MessageItemRequestBuilder = self._get_message(message_id)
        return await message.get()

    async def get_messages(
        self,
        folder_id: str,
        *,
        orderby: tuple[str, ...] = DEFAULT_ORDERBY,
        select: tuple[str, ...] = DEFAULT_SELECT,
        top: int | None = None,
        filter: str | None = None,
    ) -> tuple[list[Message], bool]:
        """Retrieve messages from a specific folder with filtering options.

        Fetches a list of messages from the specified folder with customizable
        sorting, field selection, and result limiting.

        Args:
            folder_id: The ID of the folder to retrieve messages from.
            orderby: Tuple of OData orderby clauses (e.g., "receivedDateTime DESC").
                Defaults to sorting by received date descending.
            select: Tuple of field names to include in results. Defaults to
                common fields like id, subject, from, recipients, dates, etc.
            top: Maximum number of messages to return. None for default limit.
            filter: OData filter expression for filtering messages (e.g., "receivedDateTime ge 2025-10-03T00:00:00Z").

        Returns:
            tuple[list[Message], bool]: A tuple containing:
                - List of message objects matching the query (empty list if none found)
                - Boolean indicating if more messages are available (True if pagination link exists)
        """
        folder: MailFolderItemRequestBuilder = self._get_folder(folder_id)
        messages: MessageCollectionResponse | None = await folder.messages.get(
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                    orderby=list(orderby), select=list(select), top=top, filter=filter
                )
            )
        )
        return (
            messages.value if messages and messages.value else [],
            getattr(messages, "odata_next_link", None) is not None,
        )

    async def move_messages(self, folder_id: str, message_ids: tuple[str, ...]) -> list[Message | None]:
        """Move multiple messages to a different folder in parallel.

        Relocates the specified messages to the target folder, preserving all
        message properties and content. Operations are executed in parallel
        with a maximum of 4 concurrent requests to comply with API guidelines.

        Args:
            folder_id: The ID of the destination folder.
            message_ids: Tuple of message IDs to move.

        Returns:
            tuple[Message | None, ...]: Tuple of moved message objects with updated folder IDs,
                or None for any operations that failed. Results are in the same order as input IDs.
        """
        semaphore = Semaphore(4)

        async def move_with_semaphore(message_id: str) -> Message | None:
            async with semaphore:
                return await self._get_message(message_id).move.post(
                    MovePostRequestBody(destination_id=folder_id)
                )

        return await gather(*(move_with_semaphore(message_id) for message_id in message_ids))
