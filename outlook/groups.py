import asyncio
from collections.abc import Callable
from functools import wraps
from typing import Any

from click import Command, Group


class AsyncGroup(Group):
    """Click Group with support for async commands"""

    def async_command(self, *args: Any, **kwargs: Any) -> Callable[[Callable[..., Any]], Command]:
        """Decorator for async commands that automatically wraps them with asyncio.run"""
        def decorator(f: Callable[..., Any]) -> Command:
            @wraps(f)
            def wrapper(*cmd_args: Any, **cmd_kwargs: Any) -> Any:
                return asyncio.run(f(*cmd_args, **cmd_kwargs))
            return self.command(*args, **kwargs)(wrapper)
        return decorator
