"""Main entry point: starts configured channel adapters + orchestrator.

Channels:
  - Slack: Socket Mode (async, no port)
  - Telegram: Long polling
  - Teams: Bot Framework webhook (aiohttp on configurable port)
"""

import asyncio
import logging
import os
import signal

# Claude Agent SDK spawns CLI subprocesses. If this process was launched
# from inside a Claude Code session, CLAUDECODE=1 leaks into child envs
# and causes "nested session" crashes. Strip it early.
for _key in ("CLAUDECODE", "CLAUDE_CODE_ENTRYPOINT"):
    os.environ.pop(_key, None)

from orchestrator import CONFIG
from orchestrator.server import ConfirmGate, register_channel

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(name)s] %(levelname)s: %(message)s",
)
logger = logging.getLogger(__name__)


async def main() -> None:
    confirm_gate = ConfirmGate()
    channels_config = CONFIG.get("channels", {})
    tasks: list[asyncio.Task] = []

    # --- Slack ---
    if channels_config.get("slack", {}).get("enabled"):
        from orchestrator.channel.slack import SlackChannel

        slack_ch = SlackChannel(confirm_gate)
        register_channel("slack", slack_ch)
        tasks.append(asyncio.create_task(slack_ch.start()))
        logger.info("  Slack: Socket Mode")

    # --- Telegram ---
    if channels_config.get("telegram", {}).get("enabled"):
        from orchestrator.channel.telegram import TelegramChannel

        tg_ch = TelegramChannel(confirm_gate)
        register_channel("telegram", tg_ch)
        tasks.append(asyncio.create_task(tg_ch.start()))
        logger.info("  Telegram: Long Polling")

    # --- Teams ---
    teams_cfg = channels_config.get("teams", {})
    if teams_cfg.get("enabled"):
        from orchestrator.channel.teams import TeamsChannel

        port = teams_cfg.get("port", 3978)
        teams_ch = TeamsChannel(confirm_gate, port=port)
        register_channel("teams", teams_ch)
        tasks.append(asyncio.create_task(teams_ch.start()))
        logger.info("  Teams: Bot Framework webhook on port %d", port)

    loop = asyncio.get_running_loop()
    stop_event = asyncio.Event()

    def _shutdown() -> None:
        logger.info("Shutdown signal received")
        stop_event.set()

    for sig in (signal.SIGINT, signal.SIGTERM):
        loop.add_signal_handler(sig, _shutdown)

    logger.info("Starting orchestrator...")
    await stop_event.wait()

    logger.info("Shutting down...")
    from orchestrator.server import _channels

    for name, adapter in list(_channels.items()):
        try:
            await adapter.stop()
        except Exception:
            pass
    for t in tasks:
        t.cancel()

    logger.info("Orchestrator stopped.")


if __name__ == "__main__":
    asyncio.run(main())