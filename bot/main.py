import asyncio
import logging
import os

import discord
from discord.ext import commands

from bot.config import BOT_PREFIX, BOT_TOKEN

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("bot")

# Intents
intents = discord.Intents.default()
intents.message_content = True

bot = commands.Bot(command_prefix=BOT_PREFIX, intents=intents)

# Path to cogs directory
COGS_DIR = os.path.join(os.path.dirname(__file__), "cogs")


async def load_cogs():
    """Auto-load all cog files from bot/cogs/."""
    for filename in os.listdir(COGS_DIR):
        if filename.endswith(".py") and not filename.startswith("_"):
            cog_name = f"bot.cogs.{filename[:-3]}"
            try:
                await bot.load_extension(cog_name)
                log.info("Loaded cog: %s", cog_name)
            except Exception:
                log.exception("Failed to load cog: %s", cog_name)


@bot.event
async def on_ready():
    log.info("Logged in as %s (ID: %s)", bot.user, bot.user.id)
    log.info("Connected to %d guild(s)", len(bot.guilds))


async def main():
    if not BOT_TOKEN:
        log.error("DISCORD_BOT_TOKEN is not set. Please check your .env file.")
        return

    async with bot:
        await load_cogs()
        await bot.start(BOT_TOKEN)


if __name__ == "__main__":
    asyncio.run(main())
