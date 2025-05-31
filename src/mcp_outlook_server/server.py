import asyncio
from .common import logger, mcp
from . import resources, tools

async def main():
    logger.info("Starting Outlook MCP server")
    logger.info("Running MCP server...")
    await mcp.run_stdio_async()

if __name__ == "__main__":
    asyncio.run(main())