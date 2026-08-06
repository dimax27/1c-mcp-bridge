"""
1C MCP Bridge — HTTP server for Qwen Desktop.

Qwen restricts stdio MCP to npx/uvx only. Run this over HTTP instead.

Start via Start Menu shortcut or:
    python mcp_server_1c_http.py --port 8000

Security: binds to 127.0.0.1 by default. Use --insecure-host to bind
to other interfaces (not recommended — any local process can access it).
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from mcp_server_1c import mcp, log, DB_CONFIG

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--insecure-host", action="store_true",
        help="Allow binding to non-localhost addresses (DANGER: no auth)")
    args = parser.parse_args()

    if args.host not in ("127.0.0.1", "localhost", "::1"):
        if not args.insecure_host:
            parser.error(
                f"Binding to {args.host} requires --insecure-host flag. "
                "Non-localhost binding exposes 1C data to the network without authentication."
            )

    log.info("Starting 1C MCP Bridge HTTP on %s:%d", args.host, args.port)
    log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))
    mcp.run(transport="streamable-http", host=args.host, port=args.port)
