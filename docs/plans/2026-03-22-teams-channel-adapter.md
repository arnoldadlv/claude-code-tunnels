# Teams Channel Adapter Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a Microsoft Teams channel adapter so users can @mention the bot in Teams channels to trigger orchestrated tasks.

**Architecture:** A `TeamsChannel` class extending `BaseChannel`, using the `botbuilder-integration-aiohttp` SDK. The adapter runs an aiohttp web server on a dedicated port to receive webhook POSTs from Azure Bot Service. Async replies use saved `ConversationReference` objects via `ADAPTER.continue_conversation()`.

**Tech Stack:** botbuilder-integration-aiohttp (Bot Framework Python SDK), aiohttp, existing BaseChannel/session infrastructure.

---

### Task 1: Add dependency

**Files:**
- Modify: `requirements.txt`

**Step 1: Add botbuilder package to requirements**

In `requirements.txt`, add after the existing dependencies:

```
# Teams (install if using Teams channel)
botbuilder-integration-aiohttp>=4.14.5
```

**Step 2: Commit**

```bash
git add requirements.txt
git commit -m "deps: add botbuilder-integration-aiohttp for Teams channel"
```

---

### Task 2: Create the Teams channel adapter

**Files:**
- Create: `orchestrator/channel/teams.py`

**Step 1: Write the adapter**

```python
"""Microsoft Teams channel adapter: Bot Framework webhook receive + async reply.

Uses botbuilder-integration-aiohttp to handle incoming Activities from Azure Bot
Service. Inherits BaseChannel for shared session management, confirm/cancel flow,
and message splitting.
"""

from __future__ import annotations

import asyncio
import logging
import sys
import traceback
from pathlib import Path
from typing import Any, TYPE_CHECKING

from aiohttp import web
from botbuilder.core import TurnContext, MessageFactory
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.integration.aiohttp import CloudAdapter, ConfigurationBotFrameworkAuthentication
from botbuilder.schema import Activity

from orchestrator import ARCHIVE_PATH
from orchestrator.channel.base import BaseChannel, load_credential_file, split_message

if TYPE_CHECKING:
    from orchestrator.server import ConfirmGate

logger = logging.getLogger(__name__)

CREDENTIAL_PATH = ARCHIVE_PATH / "teams" / "credentials"
DEFAULT_PORT = 3978


def load_credentials(path: Path | None = None) -> dict[str, str]:
    p = path or CREDENTIAL_PATH
    return load_credential_file(p)


class _BotFrameworkConfig:
    """Config object that ConfigurationBotFrameworkAuthentication reads from."""

    def __init__(self, app_id: str, app_password: str, app_type: str = "MultiTenant"):
        self.APP_ID = app_id
        self.APP_PASSWORD = app_password
        self.APP_TYPE = app_type
        self.APP_TENANTID = ""


class TeamsChannel(BaseChannel):
    """Teams channel: receives messages via Bot Framework webhook, sends via Bot Connector.

    Shares BaseChannel's session state machine, confirm/cancel flow, and
    message splitting. Only the transport layer (Bot Framework SDK) is
    Teams-specific.
    """

    channel_name = "teams"

    def __init__(self, confirm_gate: ConfirmGate, port: int = DEFAULT_PORT) -> None:
        super().__init__(confirm_gate)
        creds = load_credentials()
        self._app_id = creds["app_id"]
        self._port = port

        config = _BotFrameworkConfig(
            app_id=creds["app_id"],
            app_password=creds["app_password"],
            app_type=creds.get("app_type", "MultiTenant"),
        )
        self._adapter = CloudAdapter(ConfigurationBotFrameworkAuthentication(config))
        self._adapter.on_turn_error = self._on_turn_error

        allowed = creds.get("allowed_users", "")
        self._allowed_users: set[str] = set()
        if allowed:
            self._allowed_users = {u.strip() for u in allowed.split(",") if u.strip()}

        # Store conversation references for async replies, keyed by conversation ID
        self._conv_refs: dict[str, Any] = {}

        self._runner: web.AppRunner | None = None
        self._bot = _TeamsBot(self)

    # -- Transport: send -------------------------------------------------------

    async def _send(self, callback_info: Any, text: str) -> None:
        """BaseChannel calls this to deliver messages. We split + send via Bot Connector."""
        conv_id = callback_info.get("conversation_id", "")
        conv_ref = self._conv_refs.get(conv_id)
        if not conv_ref:
            logger.error("No conversation reference for %s, cannot send reply", conv_id)
            return

        chunks = split_message(text, max_len=4096)
        for chunk in chunks:
            await self._adapter.continue_conversation(
                conv_ref,
                lambda turn_ctx, c=chunk: turn_ctx.send_activity(MessageFactory.text(c)),
                self._app_id,
            )

    # -- Incoming message handling ---------------------------------------------

    async def _on_teams_message(self, turn_context: TurnContext) -> None:
        """Called by the inner _TeamsBot when an @mention message arrives."""
        activity = turn_context.activity

        # Strip @mention
        TurnContext.remove_recipient_mention(activity)
        text = (activity.text or "").strip()
        if not text:
            return

        user_id = activity.from_property.id if activity.from_property else ""
        user_name = activity.from_property.name if activity.from_property else ""

        if self._allowed_users:
            if user_id not in self._allowed_users and user_name not in self._allowed_users:
                logger.warning("Teams: unauthorized user %s (%s)", user_name, user_id)
                return

        # Save conversation reference for async replies
        conv_ref = TurnContext.get_conversation_reference(activity)
        conv_id = activity.conversation.id if activity.conversation else ""
        self._conv_refs[conv_id] = conv_ref

        callback_info = {
            "conversation_id": conv_id,
            "user_id": user_id,
            "user_name": user_name,
        }

        logger.info("Teams from %s (conv %s): %s", user_name or user_id, conv_id[:20], text[:100])
        await self._handle_text(text, conv_id, callback_info)

    # -- Error handler ---------------------------------------------------------

    async def _on_turn_error(self, turn_context: TurnContext, error: Exception) -> None:
        logger.error("Teams bot turn error: %s", error)
        traceback.print_exc(file=sys.stderr)
        try:
            await turn_context.send_activity("An error occurred processing your request.")
        except Exception:
            logger.exception("Failed to send error message to Teams")

    # -- Lifecycle -------------------------------------------------------------

    async def start(self) -> None:
        app = web.Application(middlewares=[aiohttp_error_middleware])
        app.router.add_post("/api/messages", self._handle_webhook)
        app.router.add_get("/health", self._handle_health)

        self._runner = web.AppRunner(app)
        await self._runner.setup()
        site = web.TCPSite(self._runner, "0.0.0.0", self._port)
        await site.start()
        logger.info("Teams channel started on port %d", self._port)

    async def stop(self) -> None:
        if self._runner:
            await self._runner.cleanup()
            self._runner = None
        logger.info("Teams channel stopped.")

    # -- Webhook endpoint ------------------------------------------------------

    async def _handle_webhook(self, request: web.Request) -> web.Response:
        return await self._adapter.process(request, self._bot)

    async def _handle_health(self, request: web.Request) -> web.Response:
        return web.json_response({"status": "ok", "channel": "teams", "port": self._port})


class _TeamsBot(TeamsActivityHandler):
    """Inner bot class that delegates to TeamsChannel for message handling."""

    def __init__(self, channel: TeamsChannel) -> None:
        super().__init__()
        self._channel = channel

    async def on_message_activity(self, turn_context: TurnContext) -> None:
        await self._channel._on_teams_message(turn_context)
```

**Step 2: Commit**

```bash
git add orchestrator/channel/teams.py
git commit -m "feat: add Teams channel adapter using Bot Framework SDK"
```

---

### Task 3: Wire Teams into main.py

**Files:**
- Modify: `orchestrator/main.py`

**Step 1: Add Teams startup block**

After the Telegram startup block (around line 25), add:

```python
    if channels_config.get("teams", {}).get("enabled"):
        from orchestrator.channel.teams import TeamsChannel
        teams_port = channels_config.get("teams", {}).get("port", 3978)
        teams_ch = TeamsChannel(confirm_gate, port=teams_port)
        register_channel("teams", teams_ch)
        tasks.append(asyncio.create_task(teams_ch.start()))
        logger.info("  Teams: Bot Framework webhook on port %d", teams_port)
```

**Step 2: Commit**

```bash
git add orchestrator/main.py
git commit -m "feat: wire Teams channel into orchestrator main"
```

---

### Task 4: Update orchestrator.yaml config template

**Files:**
- Modify: `orchestrator.yaml`

**Step 1: Add teams section**

Under the `channels` key, after `telegram`, add:

```yaml
  teams:
    enabled: false
    port: 3978
```

**Step 2: Commit**

```bash
git add orchestrator.yaml
git commit -m "config: add Teams channel section to orchestrator.yaml"
```

---

### Task 5: Create /connect-teams skill

**Files:**
- Create: `skills/connect-teams/SKILL.md`

**Step 1: Write the skill**

```markdown
---
name: connect-teams
description: "Connect Microsoft Teams to an existing Orchestrator. Guides through Azure Bot registration, credential input, webhook configuration, and connection test. Run with /connect-teams. Use for requests like 'connect teams', 'add teams channel'."
---

# Connect Teams Channel

Connects a Microsoft Teams channel to an already-installed Orchestrator.

## Prerequisites

- `orchestrator/` directory must exist in the current folder
- `orchestrator.yaml` must exist
- A publicly reachable HTTPS URL for the bot webhook endpoint

## Flow

### Step 1: Verify Orchestrator

1. Check `orchestrator.yaml` exists in current directory
2. Load config to find ARCHIVE_PATH
3. If not found -> "Please run /setup-orchestrator first"

### Step 2: Azure Bot Registration Guide

If no credentials found, show the user:

    Azure Bot Setup Guide:
    1. Go to https://portal.azure.com -> Create a resource -> "Azure Bot"
    2. Fill in:
       - Bot handle: choose a unique name
       - Subscription & Resource Group: select yours
       - Pricing: F0 (free) is fine for testing
       - Type of App: Multi Tenant
       - Creation type: "Create new Microsoft App ID"
    3. After creation, go to the Bot resource
    4. Settings -> Configuration:
       - Messaging endpoint: https://YOUR-PUBLIC-HOST:3978/api/messages
       - (You need a public HTTPS URL — use ngrok, Cloudflare Tunnel, or a reverse proxy)
    5. Settings -> Configuration -> "Manage Password" (next to Microsoft App ID)
       - Click "New client secret", copy the Value (this is your app_password)
       - Copy the Application (client) ID (this is your app_id)
    6. Channels -> Microsoft Teams -> Save (enables the Teams channel)
    7. Go to https://teams.microsoft.com -> Apps -> search for your bot name -> Add to a team

### Step 3: Collect Credentials

    app_id      : (Application/client ID from Azure)
    app_password : (Client secret value)
    app_type     : MultiTenant (or SingleTenant if your org requires it)
    allowed_users : (optional, comma-separated Teams user IDs or display names)

### Step 4: Save & Configure

1. Create ARCHIVE_PATH/teams/credentials with the collected values
2. Update orchestrator.yaml: set channels.teams.enabled = true
3. Optionally set channels.teams.port (default 3978)
4. Install dependencies: `pip install botbuilder-integration-aiohttp>=4.14.5`

### Step 5: Networking Check

Before starting, confirm the webhook URL is reachable:

    # If using ngrok for development:
    ngrok http 3978

    # If using Cloudflare Tunnel:
    cloudflared tunnel --url http://localhost:3978

    # The messaging endpoint in Azure Bot config must match:
    # https://YOUR-DOMAIN/api/messages

### Step 6: Test

1. Restart orchestrator (or start it): ./start-orchestrator.sh --fg
2. Check logs for "Teams channel started on port 3978"
3. In Microsoft Teams, go to a channel where the bot is added
4. @mention the bot with a test message, e.g.: @OrchestratorBot hello
5. Verify the confirm/cancel flow works

If no response:
- Check orchestrator logs for errors
- Verify the messaging endpoint URL in Azure Bot settings
- Verify the HTTPS tunnel is running
- Verify the bot is added to the Teams channel

## Credential File Format

    app_id : xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
    app_password : your-client-secret-value
    app_type : MultiTenant
    allowed_users : User One, User Two

## Rules

- If orchestrator not installed, redirect to /setup-orchestrator
- Never overwrite existing credentials without confirmation
- The webhook endpoint MUST be HTTPS — Teams/Bot Framework will not send to HTTP
- If allowed_users is empty, all users in channels where the bot is added can interact
```

**Step 2: Commit**

```bash
git add skills/connect-teams/SKILL.md
git commit -m "feat: add /connect-teams setup skill"
```

---

### Task 6: Update setup-orchestrator skill to include Teams option

**Files:**
- Modify: `skills/setup-orchestrator/SKILL.md`

**Step 1: Update channel selection in Phase 1**

In the "Collect User Input" section, change the channels prompt from:

```
3. Channels to enable: slack / telegram / multiple (required — must ask)
```

to:

```
3. Channels to enable: slack / telegram / teams / multiple (required — must ask)
```

**Step 2: Add Teams section in Phase 6 (Install Dependencies)**

After the Telegram comment, add:

```bash
$PIP_CMD install botbuilder-integration-aiohttp  # if Teams
```

**Step 3: Add Teams block in Phase 7 (Channel Setup)**

After the Telegram section, add:

```
#### Teams
1. Check if ARCHIVE_PATH/teams/credentials exists
2. If not -> show Azure Bot registration guide (same as /connect-teams Step 2)
3. Collect: app_id, app_password, app_type, optionally allowed_users
4. Create credential file
5. Remind user they need a public HTTPS URL for the messaging endpoint
```

**Step 4: Commit**

```bash
git add skills/setup-orchestrator/SKILL.md
git commit -m "feat: add Teams option to setup-orchestrator skill"
```

---

### Task 7: Final verification and PR

**Step 1: Verify all new/modified files are present**

```bash
git diff --stat main
```

Expected files:
- `requirements.txt` (modified)
- `orchestrator/channel/teams.py` (new)
- `orchestrator/main.py` (modified)
- `orchestrator.yaml` (modified)
- `skills/connect-teams/SKILL.md` (new)
- `skills/setup-orchestrator/SKILL.md` (modified)
- `docs/plans/2026-03-22-teams-channel-adapter.md` (new)

**Step 2: Verify Python syntax**

```bash
python3 -c "import ast; ast.parse(open('orchestrator/channel/teams.py').read()); print('OK')"
```

Expected: `OK`

**Step 3: Push and create PR**

```bash
git push -u origin feature/teams-channel-adapter
gh pr create --title "feat: add Microsoft Teams channel adapter" --body "..."
```
