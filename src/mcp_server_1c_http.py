"""
1C MCP Bridge — HTTP version for Qwen Desktop.

Qwen Desktop restricts stdio MCP commands to `npx`/`uvx` only.
This script runs the same MCP server over Streamable HTTP instead.

Usage:
    python mcp_server_1c_http.py [--port 8000]

In Qwen Desktop: StreamableHTTP → http://127.0.0.1:8000/mcp

The server can also be run persistently:
    - Create a scheduled task to start it on login
    - Or keep a PowerShell window open
"""
import sys
import os

# Ensure we can import mcp_server_1c from the same directory
sys.path.insert(0, os.path.dirname(__file__))
from mcp_server_1c import mcp, log, DB_CONFIG

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="1C MCP Bridge — HTTP server")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    args = parser.parse_args()

    log.info("Starting 1C MCP Bridge HTTP on %s:%d", args.host, args.port)
    log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))

    mcp.run(transport="streamable-http", host=args.host, port=args.port)
