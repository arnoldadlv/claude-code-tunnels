"""Channel adapters: Slack, Telegram, Teams, with shared base and session management."""

from orchestrator.channel.base import (
    BaseChannel,
    CONFIRM_KEYWORDS,
    CANCEL_KEYWORDS,
    load_credential_file,
    split_message,
)
from orchestrator.channel.session import (
    Session,
    SessionState,
    SessionStore,
    FOLLOWUP_END_KEYWORDS,
)

__all__ = [
    "BaseChannel",
    "CONFIRM_KEYWORDS",
    "CANCEL_KEYWORDS",
    "FOLLOWUP_END_KEYWORDS",
    "Session",
    "SessionState",
    "SessionStore",
    "load_credential_file",
    "split_message",
]
