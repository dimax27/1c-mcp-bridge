#!/usr/bin/env node
/**
 * 1C MCP Bridge — Node.js launcher for Qwen Desktop.
 *
 * Qwen Desktop only allows `npx` and `uvx` as commands for stdio MCP servers.
 * This thin wrapper spawns the Python MCP server and pipes stdin/stdout.
 *
 * Used by: Qwen Desktop MCP config
 * Not needed by: Claude Desktop, Kimi Desktop, Reasonix (they accept python.exe directly)
 */

const { spawn } = require("child_process");
const path = require("path");

const launcherDir = __dirname;
const pythonExe = process.env.ONEC_PYTHON_EXE || path.join(launcherDir, "..", ".venv", "Scripts", "python.exe");
const serverScript = process.env.ONEC_SERVER_SCRIPT || path.join(launcherDir, "mcp_server_1c.py");
const databasesFile = process.env.ONEC_DATABASES_FILE || "";

const env = { ...process.env };
if (databasesFile) {
    env.ONEC_DATABASES_FILE = databasesFile;
}

const proc = spawn(pythonExe, [serverScript], {
    stdio: "inherit",
    env: env,
});

proc.on("error", (err) => {
    console.error("Failed to launch 1C MCP Bridge:", err.message);
    process.exit(1);
});

proc.on("exit", (code) => {
    process.exit(code || 0);
});
