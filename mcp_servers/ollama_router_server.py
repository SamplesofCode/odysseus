"""
ollama_router_server.py

MCP server that lets the agent query models on remote Ollama instances.
Reads cookbook_state.json to discover servers with Ollama (probe_port 11434),
supports WoL wake-up, and exposes tools for listing models and running inference.
"""

import asyncio
import json
import sys
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

server = Server("ollama_router")

# ── Helpers ──────────────────────────────────────────────────────────────────

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
COOKBOOK_STATE = DATA_DIR / "cookbook_state.json"
OLLAMA_PORT = 11434
DEFAULT_TIMEOUT = 120  # seconds for generation


def _ollama_servers() -> list[dict]:
    """Return cookbook servers that have Ollama (probe_port == '11434')."""
    try:
        state = json.loads(COOKBOOK_STATE.read_text(encoding="utf-8"))
    except Exception:
        return []
    servers = state.get("env", {}).get("servers", [])
    results = []
    for s in servers:
        host = s.get("host", "")
        probe = str(s.get("probe_port", ""))
        if probe == str(OLLAMA_PORT) and host:
            # Extract IP from "user@ip" format
            ip = host.split("@")[-1] if "@" in host else host
            results.append({
                "name": s.get("name", ip),
                "ip": ip,
                "port": OLLAMA_PORT,
                "mac": s.get("mac", ""),
                "broadcast": s.get("broadcast", "255.255.255.255"),
            })
    return results


def _find_server(name: str | None) -> dict | None:
    """Find an Ollama server by name (case-insensitive), or return the first one."""
    servers = _ollama_servers()
    if not servers:
        return None
    if not name:
        return servers[0]
    name_lower = name.lower()
    for s in servers:
        if s["name"].lower() == name_lower:
            return s
    return None


async def _ollama_api(ip: str, port: int, endpoint: str,
                      payload: dict | None = None,
                      timeout: float = DEFAULT_TIMEOUT) -> dict | None:
    """Make an HTTP request to an Ollama API endpoint."""
    import httpx
    base = f"http://{ip}:{port}"
    url = f"{base}{endpoint}"
    try:
        async with httpx.AsyncClient(
            timeout=httpx.Timeout(connect=10.0, read=timeout, write=30.0, pool=30.0)
        ) as client:
            if payload is not None:
                resp = await client.post(url, json=payload)
            else:
                resp = await client.get(url)
            if resp.status_code == 200:
                return resp.json()
            return {"error": f"HTTP {resp.status_code}: {resp.text[:300]}"}
    except httpx.ConnectError:
        return {"error": f"Cannot connect to Ollama at {base}. Is the server online?"}
    except httpx.TimeoutException:
        return {"error": f"Request to {base} timed out after {timeout}s"}
    except Exception as e:
        return {"error": str(e)}


# ── Tools ────────────────────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="ollama_list_models",
            description=(
                "List models available on a remote Ollama server. "
                "Discovers servers from cookbook config (those with probe_port 11434). "
                "Returns model names, sizes, and quantization levels."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "server": {
                        "type": "string",
                        "description": "Server name from cookbook (e.g. 'Desktop'). Uses first Ollama server if omitted.",
                    },
                },
            },
        ),
        Tool(
            name="ollama_query",
            description=(
                "Send a prompt to a model on a remote Ollama server and get the response. "
                "Use this to delegate complex reasoning, coding, or analysis tasks to a "
                "bigger/more capable model running on a different machine. "
                "The remote model processes the prompt and returns its full response."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "prompt": {
                        "type": "string",
                        "description": "The prompt to send to the remote model.",
                    },
                    "model": {
                        "type": "string",
                        "description": "Model name on the remote Ollama (e.g. 'deepseek-r1-0528-qwen3-8b'). If omitted, uses the server's default.",
                    },
                    "system": {
                        "type": "string",
                        "description": "Optional system prompt for the query.",
                    },
                    "server": {
                        "type": "string",
                        "description": "Server name from cookbook (e.g. 'Desktop'). Uses first Ollama server if omitted.",
                    },
                    "temperature": {
                        "type": "number",
                        "description": "Sampling temperature (default: 0.7).",
                    },
                },
                "required": ["prompt"],
            },
        ),
        Tool(
            name="ollama_status",
            description=(
                "Check if a remote Ollama server is reachable and what model (if any) "
                "is currently loaded. Can optionally wake the server via WoL first."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "server": {
                        "type": "string",
                        "description": "Server name from cookbook (e.g. 'Desktop').",
                    },
                    "wake": {
                        "type": "boolean",
                        "description": "If true and the server is offline but WoL-capable, send a wake packet and wait for it to come online (up to 90s).",
                    },
                },
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    if name == "ollama_list_models":
        return await _tool_list_models(arguments)
    elif name == "ollama_query":
        return await _tool_query(arguments)
    elif name == "ollama_status":
        return await _tool_status(arguments)
    return [TextContent(type="text", text=f"Unknown tool: {name}")]


async def _tool_list_models(args: dict) -> list[TextContent]:
    srv = _find_server(args.get("server"))
    if not srv:
        return [TextContent(type="text", text="Error: No Ollama servers found in cookbook config.")]

    data = await _ollama_api(srv["ip"], srv["port"], "/api/tags")
    if not data or "error" in data:
        err = data.get("error", "Unknown error") if data else "No response"
        return [TextContent(type="text", text=f"Error reaching {srv['name']} ({srv['ip']}): {err}")]

    models = data.get("models", [])
    if not models:
        return [TextContent(type="text", text=f"{srv['name']} ({srv['ip']}) has no models pulled. Run: ollama pull <model>")]

    lines = [f"Models on {srv['name']} ({srv['ip']}:{srv['port']}):"]
    for m in models:
        name = m.get("name", "?")
        size_bytes = m.get("size", 0)
        size_gb = size_bytes / (1024 ** 3) if size_bytes else 0
        details = m.get("details", {})
        quant = details.get("quantization_level", "")
        params = details.get("parameter_size", "")
        info_parts = []
        if params:
            info_parts.append(params)
        if quant:
            info_parts.append(quant)
        if size_gb > 0:
            info_parts.append(f"{size_gb:.1f}GB")
        info = f" ({', '.join(info_parts)})" if info_parts else ""
        lines.append(f"  - {name}{info}")

    return [TextContent(type="text", text="\n".join(lines))]


async def _tool_query(args: dict) -> list[TextContent]:
    prompt = args.get("prompt", "")
    if not prompt:
        return [TextContent(type="text", text="Error: prompt is required.")]

    srv = _find_server(args.get("server"))
    if not srv:
        return [TextContent(type="text", text="Error: No Ollama servers found in cookbook config.")]

    model = args.get("model", "")
    system = args.get("system", "")
    temperature = args.get("temperature", 0.7)

    # If no model specified, try to pick the first available one
    if not model:
        tags = await _ollama_api(srv["ip"], srv["port"], "/api/tags", timeout=10)
        if tags and "models" in tags and tags["models"]:
            model = tags["models"][0]["name"]
        else:
            return [TextContent(type="text", text=f"Error: No model specified and couldn't list models on {srv['name']}.")]

    # Use the /api/chat endpoint (OpenAI-style messages)
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})

    payload = {
        "model": model,
        "messages": messages,
        "stream": False,
        "options": {"temperature": temperature},
    }

    data = await _ollama_api(srv["ip"], srv["port"], "/api/chat", payload=payload)
    if not data or "error" in data:
        err = data.get("error", "Unknown error") if data else "No response"
        return [TextContent(type="text", text=f"Error from {srv['name']}: {err}")]

    # Extract response
    message = data.get("message", {})
    content = message.get("content", "")
    if not content:
        return [TextContent(type="text", text=f"Empty response from {model} on {srv['name']}.")]

    # Include metadata
    eval_duration = data.get("eval_duration", 0)
    total_duration = data.get("total_duration", 0)
    eval_count = data.get("eval_count", 0)

    meta_parts = [f"Model: {model} on {srv['name']}"]
    if eval_count and eval_duration:
        tokens_per_sec = eval_count / (eval_duration / 1e9) if eval_duration else 0
        meta_parts.append(f"{eval_count} tokens, {tokens_per_sec:.1f} tok/s")
    if total_duration:
        meta_parts.append(f"total: {total_duration / 1e9:.1f}s")

    header = " | ".join(meta_parts)
    return [TextContent(type="text", text=f"[{header}]\n\n{content}")]


async def _tool_status(args: dict) -> list[TextContent]:
    srv = _find_server(args.get("server"))
    if not srv:
        return [TextContent(type="text", text="Error: No Ollama servers found in cookbook config.")]

    from src.wol import check_port, check_arp, send_wol, wait_for_host

    port_ok = await check_port(srv["ip"], srv["port"])

    if port_ok:
        # Check what's loaded
        data = await _ollama_api(srv["ip"], srv["port"], "/api/tags", timeout=10)
        models = data.get("models", []) if data and "error" not in data else []
        model_names = [m["name"] for m in models] if models else []
        return [TextContent(type="text", text=(
            f"{srv['name']} ({srv['ip']}:{srv['port']}): ONLINE\n"
            f"Available models: {', '.join(model_names) if model_names else 'none'}"
        ))]

    # Server is not reachable
    arp_ok = await check_arp(srv["ip"])

    if arp_ok:
        return [TextContent(type="text", text=(
            f"{srv['name']} ({srv['ip']}): Host is UP but Ollama is not responding on port {srv['port']}.\n"
            "Ollama may still be starting. Try again in a moment."
        ))]

    # Host is completely offline
    wake = args.get("wake", False)
    if wake and srv["mac"]:
        send_wol(srv["mac"], broadcast=srv["broadcast"])
        result = await wait_for_host(
            srv["ip"],
            mac=srv["mac"],
            broadcast=srv["broadcast"],
            port=srv["port"],
            timeout=90,
        )
        if result == "online":
            return [TextContent(type="text", text=f"{srv['name']} woke up successfully. Ollama is ready.")]
        elif result == "booting":
            return [TextContent(type="text", text=f"{srv['name']} is booting (host responded) but Ollama isn't ready yet. Try again shortly.")]
        else:
            return [TextContent(type="text", text=f"{srv['name']} did not respond after WoL. Check BIOS WoL settings and network.")]
    elif srv["mac"]:
        return [TextContent(type="text", text=(
            f"{srv['name']} ({srv['ip']}): OFFLINE (WoL available)\n"
            f"Call ollama_status with wake=true to send a Wake-on-LAN packet."
        ))]
    else:
        return [TextContent(type="text", text=f"{srv['name']} ({srv['ip']}): OFFLINE (no WoL configured)")]


# ── Entrypoint ───────────────────────────────────────────────────────────────

async def run():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    asyncio.run(run())
