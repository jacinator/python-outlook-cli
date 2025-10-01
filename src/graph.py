import json
from pathlib import Path
from typing import ClassVar, Final

from azure.core.credentials import AccessToken
from azure.identity import (
    InteractiveBrowserCredential,
    AuthenticationRecord,
    TokenCachePersistenceOptions,
)
from msgraph.graph_service_client import GraphServiceClient
from msgraph.generated.models.message_collection_response import (
    MessageCollectionResponse,
)
from msgraph.generated.models.user import User
from msgraph.generated.users.item.user_item_request_builder import (
    UserItemRequestBuilder,
)
from msgraph.generated.users.item.mail_folders.item.mail_folder_item_request_builder import (
    MailFolderItemRequestBuilder,
)
from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import (
    MessagesRequestBuilder,
)

ROOT: Final[Path] = Path(__file__).parent.parent


class CredentialDescriptor:
    def __init__(self, filename: str) -> None:
        self.path: Path = ROOT / filename

    def __get__(
        self, instance: object, owner: type[object] | None = None
    ) -> AuthenticationRecord | None:
        if self.path.exists():
            with self.path.open("r") as fp:
                return AuthenticationRecord.deserialize(fp.read())

    def __set__(self, instance: object, value: AuthenticationRecord) -> None:
        with self.path.open("w") as fp:
            fp.write(value.serialize())

    def __delete__(self, instance: object) -> None:
        self.path.unlink(missing_ok=True)


class Graph:
    auth: ClassVar[CredentialDescriptor] = CredentialDescriptor(".auth_record.json")

    def __init__(self) -> None:
        credentials: Path = ROOT / ".auth.json"

        with credentials.open("r") as fp:
            config = json.load(fp)

        self.client_id: str = config["clientId"]
        self.tenant_id: str = config["tenantId"]

        self._scopes: str = config["graphUserScopes"]
        self.scopes: list[str] = self._scopes.split()

        self.credential: InteractiveBrowserCredential = InteractiveBrowserCredential(
            client_id=self.client_id,
            tenant_id=self.tenant_id,
            authentication_record=self.auth,
            cache_persistence_options=TokenCachePersistenceOptions(),
        )
        self.client: GraphServiceClient = GraphServiceClient(
            credentials=self.credential,
            scopes=self.scopes,
        )

    async def get_inbox(self) -> MessageCollectionResponse | None:
        folder: MailFolderItemRequestBuilder = (
            self.client.me.mail_folders.by_mail_folder_id("inbox")
        )
        messages: MessageCollectionResponse | None = await folder.messages.get(
            MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
                query_parameters=MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
                    orderby=["receivedDateTime DESC"],
                    select=["from", "isRead", "receivedDateTime", "subject"],
                    top=25,
                )
            )
        )
        return messages

    async def get_user(self) -> User | None:
        user: User | None = await self.client.me.get(
            UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
                query_parameters=UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
                    select=["displayName", "mail", "userPrincipalName"],
                ),
            ),
        )
        return user

    async def get_user_token(self) -> str:
        if not self.auth:
            self.auth = self.credential.authenticate(scopes=self.scopes)

        token: AccessToken = self.credential.get_token(self._scopes)
        return token.token
