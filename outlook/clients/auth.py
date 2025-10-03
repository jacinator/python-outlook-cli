import json
from dataclasses import dataclass
from typing import ClassVar

from azure.identity import (
    AuthenticationRecord,
    InteractiveBrowserCredential,
    TokenCachePersistenceOptions,
)
from msgraph.graph_service_client import GraphServiceClient

from .settings import AUTH, AUTH_RECORD


class AuthenticationRecordDescriptor:
    """Descriptor for persisting Azure authentication records to disk.

    Manages reading and writing authentication records to a JSON file,
    enabling silent authentication across sessions.
    """

    def __get__(self, instance: object, owner: type[object] | None = None) -> AuthenticationRecord | None:
        """Load authentication record from disk if it exists."""
        if AUTH_RECORD.exists():
            with AUTH_RECORD.open("r") as fp:
                return AuthenticationRecord.deserialize(fp.read())

    def __set__(self, instance: object, value: AuthenticationRecord) -> None:
        """Save authentication record to disk."""
        with AUTH_RECORD.open("w") as fp:
            fp.write(value.serialize())

    def __delete__(self, instance: object) -> None:
        """Delete the authentication record file."""
        AUTH_RECORD.unlink(missing_ok=True)


@dataclass(frozen=True, kw_only=True, slots=True)
class Config:
    """Azure AD application configuration.

    Holds the client ID, tenant ID, and API scopes required for
    authenticating with Microsoft Graph API.
    """

    client_id: str
    tenant_id: str
    scopes: list[str]


class GraphAuthClient:
    """Azure AD authentication client for Microsoft Graph API.

    Manages authentication credentials, token caching, and Graph API client
    creation. Supports silent authentication via cached tokens and Windows
    Account Manager (WAM), falling back to interactive browser authentication
    when needed.
    """

    auth: ClassVar[AuthenticationRecordDescriptor] = AuthenticationRecordDescriptor()

    client: GraphServiceClient
    config: Config
    credentials: InteractiveBrowserCredential

    def __init__(self) -> None:
        """Initialize authentication client and create Graph API client.

        Loads configuration from .auth.json, sets up Azure credentials with
        token caching and WAM support, and creates an authenticated Graph
        service client.
        """
        with AUTH.open("r") as fp:
            config: dict[str, str] = json.load(fp)

        self.config = Config(
            client_id=config["clientId"],
            tenant_id=config["tenantId"],
            scopes=config["graphUserScopes"].split(),
        )
        self.credentials = InteractiveBrowserCredential(
            client_id=self.config.client_id,
            tenant_id=self.config.tenant_id,
            authentication_record=self.auth,
            cache_persistence_options=TokenCachePersistenceOptions(),
        )
        self.client = GraphServiceClient(
            credentials=self.credentials,
            scopes=self.config.scopes,
        )
