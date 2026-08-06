"""
1C MCP Bridge — HTTP server for MCP-клиентов с песочницей (Qwen, ChatGPT/Codex).

Security: binds to 127.0.0.1. Uses a random path prefix if
ONEC_HTTP_TOKEN is set (e.g. /mcp/<token> instead of /mcp).
"""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
from mcp_server_1c import DB_CONFIG, log, mcp

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--insecure-host", action="store_true")
    args = parser.parse_args()

    if args.host not in ("127.0.0.1", "localhost", "::1") and not args.insecure_host:
        parser.error(f"Binding to {args.host} requires --insecure-host flag.")

    # Random path prefix if token is set (defense-in-depth for localhost)
    token = os.environ.get("ONEC_HTTP_TOKEN", "").strip()
    path_prefix = f"/mcp/{token}" if token else "/mcp"

    log.info("Starting 1C MCP Bridge HTTP on %s:%d%s", args.host, args.port, path_prefix)
    log.info("Databases: %s", list(DB_CONFIG["databases"].keys()))
    mcp.run(
        transport="streamable-http",
        host=args.host,
        port=args.port,
        streamable_http_path=path_prefix,
    )
