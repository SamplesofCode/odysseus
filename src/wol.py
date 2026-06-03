"""Wake-on-LAN and host-probing utilities (stdlib only)."""

import asyncio
import re
import socket
import sys


_MAC_RE = re.compile(r"^([0-9A-Fa-f]{2}[:\-]){5}[0-9A-Fa-f]{2}$|^[0-9A-Fa-f]{12}$")


def _parse_mac(mac: str) -> bytes:
    """Normalise a MAC address string to 6 raw bytes."""
    mac = mac.strip().replace(":", "").replace("-", "")
    if len(mac) != 12 or not all(c in "0123456789abcdefABCDEF" for c in mac):
        raise ValueError(f"Invalid MAC address: {mac!r}")
    return bytes.fromhex(mac)


def send_wol(mac: str, broadcast: str = "255.255.255.255", port: int = 9) -> None:
    """Send a Wake-on-LAN magic packet (6×0xFF + 16×MAC) via UDP broadcast."""
    mac_bytes = _parse_mac(mac)
    payload = b"\xff" * 6 + mac_bytes * 16
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
        sock.sendto(payload, (broadcast, port))


async def check_arp(ip: str) -> bool:
    """Return True if *ip* appears in the local ARP table (Layer 2, no firewall issues)."""
    try:
        proc = await asyncio.create_subprocess_exec(
            "arp", "-a",
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.DEVNULL,
        )
        stdout, _ = await asyncio.wait_for(proc.communicate(), timeout=5)
        return ip in stdout.decode(errors="replace")
    except Exception:
        return False


async def check_port(ip: str, port: int, timeout: float = 2.0) -> bool:
    """Return True if a TCP connection to *ip*:*port* succeeds."""
    try:
        _, writer = await asyncio.wait_for(
            asyncio.open_connection(ip, port),
            timeout=timeout,
        )
        writer.close()
        await writer.wait_closed()
        return True
    except Exception:
        return False


async def wait_for_host(
    ip: str,
    *,
    mac: str | None = None,
    broadcast: str = "255.255.255.255",
    port: int | None = None,
    timeout: float = 90,
    interval: float = 5,
) -> str:
    """Poll until the host is reachable.

    Returns one of:
      ``"online"``  — port probe succeeded (service ready)
      ``"booting"`` — ARP confirms host is up but port not ready (timeout hit)
      ``"timeout"`` — host never appeared
    """
    deadline = asyncio.get_event_loop().time() + timeout
    wol_interval = 15
    last_wol = 0.0
    arp_seen = False

    while asyncio.get_event_loop().time() < deadline:
        # Re-send WoL periodically during ARP phase
        now = asyncio.get_event_loop().time()
        if mac and not arp_seen and (now - last_wol) >= wol_interval:
            send_wol(mac, broadcast=broadcast)
            last_wol = now

        if not arp_seen:
            arp_seen = await check_arp(ip)

        if arp_seen and port:
            if await check_port(ip, port):
                return "online"
        elif arp_seen and not port:
            return "online"

        await asyncio.sleep(interval)

    # Final check
    if port and await check_port(ip, port):
        return "online"
    if arp_seen:
        return "booting"
    return "timeout"
