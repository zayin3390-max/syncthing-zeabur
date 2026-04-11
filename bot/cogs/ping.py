import discord
from discord.ext import commands


class Ping(commands.Cog):
    """Basic ping/pong command to verify the bot is alive."""

    def __init__(self, bot: commands.Bot):
        self.bot = bot

    @commands.hybrid_command(name="ping", description="Check bot latency")
    async def ping(self, ctx: commands.Context):
        latency_ms = round(self.bot.latency * 1000)
        await ctx.send(f"Pong! latency: {latency_ms}ms")


async def setup(bot: commands.Bot):
    await bot.add_cog(Ping(bot))
