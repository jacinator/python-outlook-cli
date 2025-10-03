from typing import Final

from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.message import Message
from msgraph.generated.models.message_collection_response import MessageCollectionResponse
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.users.item.mail_folders.item.mail_folder_item_request_builder import MailFolderItemRequestBuilder
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
from msgraph.generated.users.item.messages.item.forward.forward_post_request_body import ForwardPostRequestBody
from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody

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
