"""
1C MCP Bridge — silent HTTP server launcher (no console window).

Run via pythonw.exe to start the server in background.
Add to Task Scheduler for auto-start on login.
To stop: Task Manager → End pythonw.exe process.
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from mcp_server_1c import mcp, log, DB_CONFIG

log.info("Starting 1C MCP Bridge HTTP on 127.0.0.1:8000 (background)")
log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))

mcp.run(transport="streamable-http", host="127.0.0.1", port=8000)
