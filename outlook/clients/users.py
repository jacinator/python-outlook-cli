from dataclasses import dataclass

from msgraph.generated.users.item.user_item_request_builder import UserItemRequestBuilder
from msgraph.graph_service_client import GraphServiceClient


@dataclass(frozen=True, slots=True)
class User:
    """Store user profile information.

    Caches the current user's display name and email address to avoid
    repeated API calls.
    """

    name: str
    addr: str

    @classmethod
    async def create(cls, client: GraphServiceClient) -> "User":
        """Load user profile information and return a User instance.

        Retrieves the user's display name and email address from Microsoft
        Graph. For Work/school accounts, email is in the mail property; for
        personal accounts, it's in userPrincipalName.
        """
        user = await client.me.get(
            UserItemRequestBuilder.UserItemRequestBuilderGetRequestConfiguration(
                query_parameters=UserItemRequestBuilder.UserItemRequestBuilderGetQueryParameters(
                    select=["displayName", "mail", "userPrincipalName"]
                )
            )
        )
        return cls(
            name=getattr(user, "display_name", None) or "",
            addr=getattr(user, "mail", None) or getattr(user, "user_principal_name", None) or "",
        )
