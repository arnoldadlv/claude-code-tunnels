"""Microsoft Teams channel adapter: Bot Framework webhook receive + Bot API send.

Inherits BaseChannel for shared session management, confirm/cancel flow,
and message splitting. Only implements Teams-specific transport via
botbuilder-integration-aiohttp.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Any, TYPE_CHECKING

from aiohttp import web
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext
from botbuilder.schema import Activity

from orchestrator import ARCHIVE_PATH
from orchestrator.channel.base import BaseChannel, load_credential_file, split_message

if TYPE_CHECKING:
    from orchestrator.server import ConfirmGate

logger = logging.getLogger(__name__)

CREDENTIAL_PATH = ARCHIVE_PATH / "teams" / "credentials"


def load_credentials(path: Path | None = None) -> dict[str, str]:
    p = path or CREDENTIAL_PATH
    return load_credential_file(p)


class TeamsChannel(BaseChannel):
    """Teams channel: receives messages via Bot Framework webhook, sends via Bot API.

    Shares BaseChannel's session state machine, confirm/cancel flow, and
    message splitting. Only the transport layer (Bot Framework) is Teams-specific.
    """

    channel_name = "teams"

    def __init__(self, confirm_gate: ConfirmGate, port: int = 3978) -> None:
        super().__init__(confirm_gate)
        creds = load_credentials()
        self._app_id = creds["app_id"]
        self._app_password = creds["app_password"]
        self._port = port

        settings = BotFrameworkAdapterSettings(
            app_id=self._app_id,
            app_password=self._app_password,
        )
        self._adapter = BotFrameworkAdapter(settings)

        self._app: web.Application | None = None
        self._runner: web.AppRunner | None = None

        # Store conversation references for sending proactive messages
        self._conversation_refs: dict[str, Any] = {}

    # ── Transport: receive ───────────────────────────────────────────

    async def _handle_incoming(self, turn: TurnContext) -> None:
        """Process an incoming Teams message via Bot Framework TurnContext."""
        activity = turn.activity
        if activity.type != "message" or not activity.text:
            return

        # Store conversation reference for async replies
        ref = TurnContext.get_conversation_reference(activity)
        source_key = ref.conversation.id
        self._conversation_refs[source_key] = ref

        # Strip @mention of the bot from the text
        text = activity.text
        if activity.entities:
            for entity in activity.entities:
                if entity.type == "mention" and hasattr(entity, "mentioned"):
                    mention_name = entity.mentioned.name or ""
                    text = re.sub(
                        rf"<at>{re.escape(mention_name)}</at>\s*",
                        "",
                        text,
                    ).strip()
        text = text.strip()
        if not text:
            return

        user_id = activity.from_property.id if activity.from_property else "unknown"
        callback_info = {
            "conversation_ref": ref,
            "source_key": source_key,
            "user_id": user_id,
        }

        logger.info("Teams from %s (conv %s): %s", user_id, source_key, text[:100])
        await self._handle_text(text, source_key, callback_info)

    # ── Transport: send ──────────────────────────────────────────────

    async def _send(self, callback_info: Any, text: str) -> None:
        """BaseChannel calls this to deliver messages. We split + send via Bot API."""
        ref = callback_info["conversation_ref"]
        chunks = split_message(text, max_len=4096)
        for chunk in chunks:
            await self._send_message(ref, chunk)

    async def _send_message(self, conversation_ref: Any, text: str) -> None:
        """Send a message using continue_conversation (proactive messaging)."""
        async def _callback(turn: TurnContext) -> None:
            await turn.send_activity(Activity(type="message", text=text))

        try:
            await self._adapter.continue_conversation(
                conversation_ref, _callback, self._app_id
            )
        except Exception:
            logger.exception("Teams sendMessage failed")

    # ── HTTP webhook ─────────────────────────────────────────────────

    async def _messages_handler(self, request: web.Request) -> web.Response:
        """aiohttp handler for POST /api/messages (Bot Framework webhook)."""
        if request.content_type != "application/json":
            return web.Response(status=415)

        body = await request.json()
        activity = Activity().deserialize(body)

        auth_header = request.headers.get("Authorization", "")

        async def _on_turn(turn: TurnContext) -> None:
            await self._handle_incoming(turn)

        try:
            await self._adapter.process_activity(activity, auth_header, _on_turn)
            return web.Response(status=200)
        except Exception:
            logger.exception("Teams process_activity failed")
            return web.Response(status=500)

    # ── Lifecycle ────────────────────────────────────────────────────

    async def start(self) -> None:
        self._app = web.Application()
        self._app.router.add_post("/api/messages", self._messages_handler)

        self._runner = web.AppRunner(self._app)
        await self._runner.setup()
        site = web.TCPSite(self._runner, "0.0.0.0", self._port)
        await site.start()
        logger.info("Teams channel started on port %d", self._port)

    async def stop(self) -> None:
        if self._runner:
            await self._runner.cleanup()
        logger.info("Teams channel stopped.")
