from asyncio import Task, create_task
from dataclasses import dataclass

from msgraph.graph_service_client import GraphServiceClient

from .auth import GraphAuthClient
from .folders import Folders
from .users import User


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
    _folders_task: Task[Folders]
    _user_task: Task[User]

    def __init__(self) -> None:
        self.auth = GraphAuthClient()
        self._folders_task = create_task(Folders.create(self.client))
        self._user_task = create_task(User.create(self.client))

    def authenticate(self) -> None:
        """Perform interactive browser authentication and save credentials.

        Opens a browser window for the user to authenticate with Microsoft
        and saves the authentication record for future silent authentication.
        """
        self.auth.auth = self.auth.credentials.authenticate(scopes=self.auth.config.scopes)

    async def folders(self) -> Folders:
        """Get folders data, waiting for background load if needed."""
        return await self._folders_task

    async def user(self) -> User:
        """Get user data, waiting for background load if needed."""
        return await self._user_task

    @property
    def client(self) -> GraphServiceClient:
        return self.auth.client
