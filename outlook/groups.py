import asyncio
from collections.abc import Callable
from functools import update_wrapper, wraps
from typing import Any

from click import Command, Context, Group


class AsyncGroup(Group):
    """Click Group with support for async commands and async group callbacks"""

    def invoke(self, ctx: Context) -> Any:
        # Wrap async callback before invocation
        if self.callback and asyncio.iscoroutinefunction(self.callback):
            original_callback = self.callback
            def sync_callback(*cb_args: Any, **cb_kwargs: Any) -> Any:
                return asyncio.run(original_callback(*cb_args, **cb_kwargs))
            self.callback = update_wrapper(sync_callback, original_callback)
        return super().invoke(ctx)

    def async_command(self, *args: Any, **kwargs: Any) -> Callable[[Callable[..., Any]], Command]:
        """Decorator for async commands that automatically wraps them with asyncio.run"""
        def decorator(f: Callable[..., Any]) -> Command:
            @wraps(f)
            def wrapper(*cmd_args: Any, **cmd_kwargs: Any) -> Any:
                return asyncio.run(f(*cmd_args, **cmd_kwargs))
            return self.command(*args, **kwargs)(wrapper)
        return decorator
