import asyncio
from typing import Final

from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.mail_folder_collection_response import MailFolderCollectionResponse
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.models.user import User
from msgraph.generated.users.item.mail_folders.item.mail_folder_item_request_builder import MailFolderItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.messages.item.forward.forward_post_request_body import ForwardPostRequestBody
from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody
from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder

from .clients import Client

DEFAULT_ORDERBY: Final[list[str]] = ["receivedDateTime DESC"]
DEFAULT_SELECT: Final[list[str]] = [
    "id",
    "subject",
    "from",
    "toRecipients",
    "ccRecipients",
    "isRead",
    "receivedDateTime",
    "sentDateTime",
    "hasAttachments",
    "importance",
    "conversationId",
    "parentFolderId",
    "webLink",
    "bodyPreview",
]


def authenticate() -> None:
    client = Client()
    client.authenticate()


@Client.decorator
async def get_folders(client: GraphServiceClient, parent_folder_id: str | None = None) -> MailFolderCollectionResponse | None:
    if parent_folder_id:
        folder: MailFolderItemRequestBuilder = client.me.mail_folders.by_mail_folder_id(parent_folder_id)
        return await folder.child_folders.get()
    return await client.me.mail_folders.get()


@Client.decorator
async def get_message(client: GraphServiceClient, message_id: str) -> Message | None:
    message: MessageItemRequestBuilder = client.me.messages.by_message_id(message_id)
    return await message.get()


@Client.decorator
async def get_messages(client: GraphServiceClient, folder_id: str, *, filter: str | None = None, select: list[str] | None = None, top: int = 25) -> MessageCollectionResponse | None:
    folder: MailFolderItemRequestBuilder = client.me.mail_folders.by_mail_folder_id(folder_id)
    return await folder.messages.get(
        MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                filter=filter,
                orderby=DEFAULT_ORDERBY,
                select=DEFAULT_SELECT if select is None else select,
                top=top,
            )
        )
    )


@Client.decorator
async def get_user(client: GraphServiceClient) -> User | None:
    return await client.me.get(
        UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
            query_parameters=UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
                select=["displayName", "mail", "userPrincipalName"]
            )
        )
    )


@Client.decorator
async def message_delete(client: GraphServiceClient, message_id: str) -> None:
    message: MessageItemRequestBuilder = client.me.messages.by_message_id(message_id)
    await message.delete()


@Client.decorator
async def message_forward(client: GraphServiceClient, message_id: str, recipients: list[str], comment: str | None = None) -> None:
    message: MessageItemRequestBuilder = client.me.messages.by_message_id(message_id)
    await message.forward.post(
        ForwardPostRequestBody(
            to_recipients=[
                Recipient(email_address=EmailAddress(address=email))
                for email in recipients
            ],
            comment=comment,
        )
    )


@Client.decorator
async def message_move(client: GraphServiceClient, message_id: str, destination_folder_id: str) -> Message | None:
    message: MessageItemRequestBuilder = client.me.messages.by_message_id(message_id)
    return await message.move.post(MovePostRequestBody(destination_id=destination_folder_id))


@Client.decorator
async def message_move_all(client: GraphServiceClient, source_folder_id: str, destination_folder_id: str) -> int:
    """Move all messages from source folder to destination folder.

    Returns the count of successfully moved messages.
    Uses concurrent operations with a limit of 4 to respect API throttling.
    """
    # Collect all message IDs from the source folder
    message_ids: list[str] = []
    folder: MailFolderItemRequestBuilder = client.me.mail_folders.by_mail_folder_id(source_folder_id)

    # Fetch all messages using pagination
    next_link: str | None = None
    while True:
        if next_link:
            # Use next_link for pagination
            response = await client.me.mail_folders.by_mail_folder_id(source_folder_id).messages.with_url(next_link).get()
        else:
            response = await folder.messages.get(
                MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                    query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                        select=["id"],
                        top=999,  # Maximum page size
                    )
                )
            )

        if response and response.value:
            message_ids.extend([msg.id for msg in response.value if msg.id])
            next_link = response.odata_next_link
            if not next_link:
                break
        else:
            break

    if not message_ids:
        return 0

    # Move messages concurrently with a semaphore to limit concurrency
    semaphore = asyncio.Semaphore(4)  # Limit to 4 concurrent operations

    async def move_with_limit(msg_id: str) -> bool:
        async with semaphore:
            try:
                message: MessageItemRequestBuilder = client.me.messages.by_message_id(msg_id)
                await message.move.post(MovePostRequestBody(destination_id=destination_folder_id))
                return True
            except Exception:
                return False

    # Execute all moves concurrently
    results = await asyncio.gather(*[move_with_limit(msg_id) for msg_id in message_ids])

    return sum(results)
