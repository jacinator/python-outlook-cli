import json
from collections.abc import Callable
from pathlib import Path
from typing import Any, ClassVar, Concatenate, Final

from azure.identity import (
    InteractiveBrowserCredential,
    AuthenticationRecord,
    TokenCachePersistenceOptions,
)
from msgraph.graph_service_client import GraphServiceClient

ROOT: Final[Path] = Path(__file__).parent.parent.parent

AUTH: Final[Path] = ROOT / ".auth.json"
AUTH_RECORD: Final[Path] = ROOT / ".auth_record.json"


class CredentialDescriptor:
    def __get__(self, instance: object, owner: type[object] | None = None) -> AuthenticationRecord | None:
        if AUTH_RECORD.exists():
            with AUTH_RECORD.open("r") as fp:
                return AuthenticationRecord.deserialize(fp.read())

    def __set__(self, instance: object, value: AuthenticationRecord) -> None:
        with AUTH_RECORD.open("w") as fp:
            fp.write(value.serialize())

    def __delete__(self, instance: object) -> None:
        AUTH_RECORD.unlink(missing_ok=True)


class Client:
    auth: ClassVar[CredentialDescriptor] = CredentialDescriptor()

    client_id: str
    tenant_id: str
    _scopes: str
    scopes: list[str]

    client: GraphServiceClient
    credential: InteractiveBrowserCredential

    @classmethod
    def decorator[**P, R](cls, func: Callable[Concatenate[GraphServiceClient, P], R]) -> Callable[P, R]:
        def inner(*args: P.args, **kwargs: P.kwargs) -> R:
            with cls() as client:
                return func(client, *args, **kwargs)
        return inner

    def __init__(self) -> None:
        with AUTH.open("r") as fp:
            config = json.load(fp)

        self.client_id = config["clientId"]
        self.tenant_id = config["tenantId"]

        self._scopes = config["graphUserScopes"]
        self.scopes = self._scopes.split()

        self.credential = InteractiveBrowserCredential(
            client_id=self.client_id,
            tenant_id=self.tenant_id,
            authentication_record=self.auth,
            cache_persistence_options=TokenCachePersistenceOptions(),
        )
        self.client = GraphServiceClient(
            credentials=self.credential,
            scopes=self.scopes,
        )

    def __enter__(self) -> GraphServiceClient:
        return self.client

    def __exit__(self, *_: Any) -> None:
        pass

    def authenticate(self) -> None:
        self.auth = self.credential.authenticate(scopes=self.scopes)
