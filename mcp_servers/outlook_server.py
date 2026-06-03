"""
outlook_server.py

MCP server for Microsoft Outlook / Hotmail email via Microsoft Graph API.
Bypasses IMAP/SMTP entirely — uses OAuth2 + REST, the only auth path
Microsoft still supports for consumer accounts (@hotmail, @outlook, @live).

First run opens a browser window for interactive login. After that, the token
is cached locally and refreshes automatically without prompting.

Register in Odysseus as an MCP server:
  Command: python
  Args:    mcp_servers/outlook_server.py
  Env:     OUTLOOK_CLIENT_ID=<your Azure app client id>
           OUTLOOK_TOKEN_CACHE=<path to token cache file>  (optional)
           OUTLOOK_EMAIL=<your email>                       (optional, for identity)

Azure App Registration setup:
  1. Go to https://portal.azure.com → App registrations → New registration
  2. Name: "Odysseus Mail" (or anything)
  3. Supported account types: "Personal Microsoft accounts only"
     (or "Accounts in any org directory and personal Microsoft accounts")
  4. Redirect URI: select "Mobile and desktop applications" →
     http://localhost
     (also add https://login.microsoftonline.com/common/oauth2/nativeclient
      as a fallback)
  5. Under API permissions, add:
     - Microsoft Graph → Delegated → Mail.Read
     - Microsoft Graph → Delegated → Mail.ReadWrite
     - Microsoft Graph → Delegated → Mail.Send
     - Microsoft Graph → Delegated → User.Read
  6. Under Authentication, enable "Allow public client flows" = Yes
  7. Copy the Application (client) ID into OUTLOOK_CLIENT_ID
"""

import asyncio
import json
import logging
import os
import sys
import time
from pathlib import Path

# MSAL handles OAuth2 device code flow + token caching + refresh
try:
    import msal
except ImportError:
    print(
        "ERROR: The 'msal' package is required.\n"
        "Install it with:  pip install msal\n",
        file=sys.stderr,
    )
    sys.exit(1)

try:
    import httpx
except ImportError:
    print(
        "ERROR: The 'httpx' package is required.\n"
        "Install it with:  pip install httpx\n",
        file=sys.stderr,
    )
    sys.exit(1)

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

logging.basicConfig(level=logging.INFO, stream=sys.stderr)
logger = logging.getLogger("outlook_server")

server = Server("outlook")

# ── Configuration ──────────────────────────────────────────────────────────

CLIENT_ID = os.environ.get("OUTLOOK_CLIENT_ID", "")
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = [
    "User.Read",
    "Mail.Read",
    "Mail.ReadWrite",
    "Mail.Send",
    "MailboxSettings.ReadWrite",
]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Token cache lives next to this script by default
_DEFAULT_CACHE = str(Path(__file__).resolve().parent.parent / "data" / ".outlook_token_cache.json")
TOKEN_CACHE_PATH = os.environ.get("OUTLOOK_TOKEN_CACHE", _DEFAULT_CACHE)

MAX_BODY_CHARS = 8000


# ── MSAL / Auth ────────────────────────────────────────────────────────────

_token_cache = msal.SerializableTokenCache()
_msal_app = None


def _load_token_cache():
    """Load persisted token cache from disk."""
    if os.path.exists(TOKEN_CACHE_PATH):
        try:
            with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                _token_cache.deserialize(f.read())
            logger.info("Loaded token cache from %s", TOKEN_CACHE_PATH)
        except Exception as e:
            logger.warning("Failed to load token cache: %s", e)


def _save_token_cache():
    """Persist token cache to disk if it changed."""
    if _token_cache.has_state_changed:
        os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(_token_cache.serialize())
        logger.info("Saved token cache to %s", TOKEN_CACHE_PATH)


def _get_msal_app():
    """Return the MSAL PublicClientApplication (singleton)."""
    global _msal_app
    if _msal_app is None:
        if not CLIENT_ID:
            raise RuntimeError(
                "OUTLOOK_CLIENT_ID is not set. Create an Azure App Registration "
                "and set the env var to the Application (client) ID. See the "
                "docstring at the top of outlook_server.py for setup steps."
            )
        _load_token_cache()
        _msal_app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            token_cache=_token_cache,
        )
    return _msal_app


def _acquire_token() -> str:
    """Get a valid access token, refreshing silently if possible.
    On first run, opens a browser window for interactive login.
    Returns the access token string."""
    app = _get_msal_app()

    # Try silent acquisition first (cached / refresh token)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_token_cache()
            return result["access_token"]

    # No cached token — open browser for interactive login
    logger.info("No cached token found — opening browser for authentication...")
    result = app.acquire_token_interactive(scopes=SCOPES)
    if "access_token" not in result:
        error = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Authentication failed: {error}")

    _save_token_cache()
    logger.info("Authentication successful")
    return result["access_token"]


def _graph_headers() -> dict:
    """Return headers with a valid Bearer token for Graph API calls."""
    token = _acquire_token()
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


# ── Graph API helpers ──────────────────────────────────────────────────────

def _graph_get(path: str, params: dict = None) -> dict:
    """GET from Microsoft Graph. Raises on HTTP errors."""
    url = f"{GRAPH_BASE}{path}"
    with httpx.Client(timeout=30) as client:
        resp = client.get(url, headers=_graph_headers(), params=params)
        if resp.status_code == 401:
            # Token might have expired mid-request; clear cache and retry once
            _token_cache.has_state_changed = True
            resp = client.get(url, headers=_graph_headers(), params=params)
        resp.raise_for_status()
        return resp.json()


def _graph_post(path: str, body: dict) -> dict:
    """POST to Microsoft Graph."""
    url = f"{GRAPH_BASE}{path}"
    with httpx.Client(timeout=30) as client:
        resp = client.post(url, headers=_graph_headers(), json=body)
        if resp.status_code == 401:
            resp = client.post(url, headers=_graph_headers(), json=body)
        resp.raise_for_status()
        # sendMail returns 202 with no body
        if resp.status_code == 202 or not resp.content:
            return {"status": "ok"}
        return resp.json()


def _graph_patch(path: str, body: dict) -> dict:
    """PATCH a Graph resource."""
    url = f"{GRAPH_BASE}{path}"
    with httpx.Client(timeout=30) as client:
        resp = client.patch(url, headers=_graph_headers(), json=body)
        if resp.status_code == 401:
            resp = client.patch(url, headers=_graph_headers(), json=body)
        resp.raise_for_status()
        if not resp.content:
            return {"status": "ok"}
        return resp.json()


def _graph_delete(path: str) -> bool:
    """DELETE a Graph resource."""
    url = f"{GRAPH_BASE}{path}"
    with httpx.Client(timeout=30) as client:
        resp = client.delete(url, headers=_graph_headers())
        if resp.status_code == 401:
            resp = client.delete(url, headers=_graph_headers())
        return resp.status_code in (200, 204)


# ── Email body extraction ─────────────────────────────────────────────────

def _extract_body(message: dict) -> str:
    """Extract readable text from a Graph message object."""
    body = message.get("body", {})
    content = body.get("content", "")
    if body.get("contentType") == "html":
        # Simple HTML stripping — good enough for email bodies
        import re
        text = re.sub(r"<br\s*/?>", "\n", content, flags=re.IGNORECASE)
        text = re.sub(r"<[^>]+>", "", text)
        try:
            import html
            text = html.unescape(text)
        except ImportError:
            pass
        return text.strip()[:MAX_BODY_CHARS]
    return content.strip()[:MAX_BODY_CHARS]


def _format_recipients(recipients: list) -> str:
    """Format a list of Graph recipient objects into a readable string."""
    parts = []
    for r in (recipients or []):
        addr = r.get("emailAddress", {})
        name = addr.get("name", "")
        email = addr.get("address", "")
        if name and name != email:
            parts.append(f"{name} <{email}>")
        else:
            parts.append(email)
    return ", ".join(parts)


# ── Tool implementations ──────────────────────────────────────────────────

def _list_messages(folder: str = "inbox", max_results: int = 20,
                   unread_only: bool = False, filter_str: str = None,
                   compact: bool = False,
                   include_attachment_names: bool = False) -> list:
    """List messages from a mail folder via Graph.
    compact=True returns only id, subject, and sender — much smaller for bulk sorting.
    include_attachment_names=True adds attachment filenames (no content downloaded)."""
    if compact:
        select = "id,subject,from,hasAttachments"
    else:
        select = "id,subject,from,receivedDateTime,isRead,hasAttachments,bodyPreview"
    params = {
        "$top": min(max_results, 100),
        "$orderby": "receivedDateTime desc",
        "$select": select,
    }
    if include_attachment_names:
        params["$expand"] = "attachments($select=name,contentType,size)"
    if unread_only:
        params["$filter"] = "isRead eq false"
    elif filter_str:
        params["$filter"] = filter_str

    path = f"/me/mailFolders/{folder}/messages"
    data = _graph_get(path, params)
    results = []
    for msg in data.get("value", []):
        from_info = msg.get("from", {}).get("emailAddress", {})
        entry = {
            "message_id": msg["id"],
            "subject": msg.get("subject", "(no subject)"),
            "from": from_info.get("name", ""),
            "from_address": from_info.get("address", ""),
        }
        if not compact:
            entry["date"] = msg.get("receivedDateTime", "")
            entry["is_read"] = msg.get("isRead", False)
            entry["has_attachments"] = msg.get("hasAttachments", False)
            entry["preview"] = msg.get("bodyPreview", "")[:200]
        if include_attachment_names:
            attachments = msg.get("attachments", [])
            entry["attachment_names"] = [a.get("name", "") for a in attachments if a.get("name")]
        elif msg.get("hasAttachments"):
            entry["has_attachments"] = True
        results.append(entry)
    return results


def _read_message(message_id: str) -> dict:
    """Read a single message by ID."""
    params = {
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
                   "isRead,hasAttachments,body,bodyPreview,internetMessageId,"
                   "replyTo,conversationId",
    }
    msg = _graph_get(f"/me/messages/{message_id}", params)
    from_info = msg.get("from", {}).get("emailAddress", {})
    return {
        "message_id": msg["id"],
        "internet_message_id": msg.get("internetMessageId", ""),
        "conversation_id": msg.get("conversationId", ""),
        "subject": msg.get("subject", "(no subject)"),
        "from": from_info.get("name", ""),
        "from_address": from_info.get("address", ""),
        "to": _format_recipients(msg.get("toRecipients", [])),
        "cc": _format_recipients(msg.get("ccRecipients", [])),
        "date": msg.get("receivedDateTime", ""),
        "is_read": msg.get("isRead", False),
        "has_attachments": msg.get("hasAttachments", False),
        "body": _extract_body(msg),
    }


def _list_attachments(message_id: str) -> list:
    """List attachments on a message."""
    data = _graph_get(f"/me/messages/{message_id}/attachments")
    results = []
    for att in data.get("value", []):
        results.append({
            "id": att["id"],
            "name": att.get("name", "unnamed"),
            "content_type": att.get("contentType", ""),
            "size": att.get("size", 0),
        })
    return results


def _send_message(to: str, subject: str, body: str,
                  cc: str = None, reply_to_id: str = None) -> dict:
    """Send an email via Graph API."""
    to_list = [addr.strip() for addr in to.split(",") if addr.strip()]
    message = {
        "subject": subject,
        "body": {
            "contentType": "text",
            "content": body,
        },
        "toRecipients": [
            {"emailAddress": {"address": addr}} for addr in to_list
        ],
    }
    if cc:
        cc_list = [addr.strip() for addr in cc.split(",") if addr.strip()]
        message["ccRecipients"] = [
            {"emailAddress": {"address": addr}} for addr in cc_list
        ]

    return _graph_post("/me/sendMail", {"message": message, "saveToSentItems": True})


def _reply_to_message(message_id: str, body: str, reply_all: bool = False) -> dict:
    """Reply to an existing message."""
    endpoint = "replyAll" if reply_all else "reply"
    return _graph_post(
        f"/me/messages/{message_id}/{endpoint}",
        {"comment": body},
    )


def _search_messages(query: str, max_results: int = 20) -> list:
    """Search messages using Graph $search (KQL)."""
    params = {
        "$search": f'"{query}"',
        "$top": min(max_results, 50),
        "$select": "id,subject,from,receivedDateTime,isRead,bodyPreview,parentFolderId",
    }
    data = _graph_get("/me/messages", params)
    results = []
    for msg in data.get("value", []):
        from_info = msg.get("from", {}).get("emailAddress", {})
        results.append({
            "message_id": msg["id"],
            "subject": msg.get("subject", "(no subject)"),
            "from": from_info.get("name", ""),
            "from_address": from_info.get("address", ""),
            "date": msg.get("receivedDateTime", ""),
            "is_read": msg.get("isRead", False),
            "preview": msg.get("bodyPreview", "")[:200],
        })
    return results


def _list_folders() -> list:
    """List mail folders."""
    data = _graph_get("/me/mailFolders", {"$top": 50})
    results = []
    for folder in data.get("value", []):
        results.append({
            "id": folder["id"],
            "name": folder.get("displayName", ""),
            "unread_count": folder.get("unreadItemCount", 0),
            "total_count": folder.get("totalItemCount", 0),
        })
    return results


def _mark_read(message_id: str, read: bool = True) -> dict:
    """Mark a message as read or unread."""
    return _graph_patch(f"/me/messages/{message_id}", {"isRead": read})


def _move_message(message_id: str, destination_folder: str) -> dict:
    """Move a message to a different folder by folder name or ID."""
    # Try to resolve folder name to ID
    folders = _list_folders()
    folder_id = destination_folder
    for f in folders:
        if f["name"].lower() == destination_folder.lower():
            folder_id = f["id"]
            break

    return _graph_post(
        f"/me/messages/{message_id}/move",
        {"destinationId": folder_id},
    )


def _delete_message(message_id: str, permanent: bool = False) -> bool:
    """Delete a message. Moves to Deleted Items by default."""
    if permanent:
        return _graph_delete(f"/me/messages/{message_id}")
    _move_message(message_id, "deleteditems")
    return True


# ── Categories ─────────────────────────────────────────────────────────────

def _list_categories() -> list:
    """List the user's Outlook categories."""
    data = _graph_get("/me/outlook/masterCategories")
    return [
        {"id": c["id"], "name": c.get("displayName", ""), "color": c.get("color", "")}
        for c in data.get("value", [])
    ]


def _categorize_message(message_id: str, categories: list) -> dict:
    """Set categories on a message. Pass an empty list to clear."""
    return _graph_patch(f"/me/messages/{message_id}", {"categories": categories})


# ── Folder management ──────────────────────────────────────────────────────

def _create_folder(name: str, parent_folder_id: str = None) -> dict:
    """Create a new mail folder. If parent_folder_id is given, creates a subfolder."""
    if parent_folder_id:
        data = _graph_post(f"/me/mailFolders/{parent_folder_id}/childFolders", {"displayName": name})
    else:
        data = _graph_post("/me/mailFolders", {"displayName": name})
    return {"id": data.get("id", ""), "name": data.get("displayName", name)}


def _rename_folder(folder_id: str, new_name: str) -> dict:
    """Rename a mail folder."""
    return _graph_patch(f"/me/mailFolders/{folder_id}", {"displayName": new_name})


def _delete_folder(folder_id: str) -> bool:
    """Delete a mail folder."""
    return _graph_delete(f"/me/mailFolders/{folder_id}")


def _resolve_folder_id(name_or_id: str) -> str:
    """Resolve a folder name to its ID. Returns the input if already an ID."""
    folders = _list_folders()
    for f in folders:
        if f["name"].lower() == name_or_id.lower():
            return f["id"]
    return name_or_id


def _bulk_move_messages(message_ids: list, destination_folder: str) -> dict:
    """Move multiple messages to a folder in one operation.
    Returns counts of successful and failed moves."""
    folder_id = _resolve_folder_id(destination_folder)
    moved = 0
    failed = 0
    errors = []
    for msg_id in message_ids:
        try:
            _graph_post(
                f"/me/messages/{msg_id}/move",
                {"destinationId": folder_id},
            )
            moved += 1
        except Exception as e:
            failed += 1
            errors.append(f"{msg_id[:20]}...: {e}")
            if len(errors) > 5:
                errors.append(f"... and {len(message_ids) - moved - failed} more to process")
                break
    return {"moved": moved, "failed": failed, "errors": errors[:6]}


# ── Auto-organize ─────────────────────────────────────────────────────────

# Sender domain → folder mapping. Checked first (most specific).
# ── Dynamic rules cache ────────────────────────────────────────────────────
import time as _time

_dynamic_rules_cache: dict[str, str] | None = None
_dynamic_rules_ts: float = 0.0
_CACHE_TTL = 300  # 5 minutes


def _invalidate_rules_cache():
    """Clear the dynamic rules cache so next classify call rebuilds it."""
    global _dynamic_rules_cache, _dynamic_rules_ts
    _dynamic_rules_cache = None
    _dynamic_rules_ts = 0.0


def _get_dynamic_sender_rules() -> dict[str, str]:
    """Build a domain→folder map from live Outlook server-side rules.
    Cached for 5 minutes. Falls through to _SENDER_RULES_FALLBACK."""
    global _dynamic_rules_cache, _dynamic_rules_ts
    now = _time.time()
    if _dynamic_rules_cache is not None and (now - _dynamic_rules_ts) < _CACHE_TTL:
        return _dynamic_rules_cache

    try:
        rules = _list_rules()
        folders = {f["id"]: f["name"] for f in _list_folders()}
        result: dict[str, str] = {}
        for rule in rules:
            if not rule.get("is_enabled"):
                continue
            conds = rule.get("conditions", {})
            folder_id = rule.get("actions", {}).get("moveToFolder")
            if not folder_id:
                continue
            folder_name = folders.get(folder_id, "")
            if not folder_name:
                continue
            for domain in conds.get("senderContains", []):
                result[domain.lower().strip()] = folder_name
        _dynamic_rules_cache = result
        _dynamic_rules_ts = now
        logger.debug(f"Refreshed dynamic rules cache: {len(result)} entries")
        return result
    except Exception as e:
        logger.warning(f"Failed to load dynamic rules, using fallback: {e}")
        return {}


# ── Fallback sender rules (used when dynamic lookup unavailable) ──────────

_SENDER_RULES_FALLBACK: dict[str, str] = {
    # Financial
    "creditkarma.com": "Financial",
    "webull.com": "Financial",
    "robinhood.com": "Financial",
    "usaa.com": "Financial",
    "chase.com": "Financial",
    "bankofamerica.com": "Financial",
    "wellsfargo.com": "Financial",
    "paypal.com": "Financial",
    "venmo.com": "Financial",
    "mint.com": "Financial",
    "schwab.com": "Financial",
    "fidelity.com": "Financial",
    "vanguard.com": "Financial",
    "sofi.com": "Financial",
    "capitalone.com": "Financial",
    "discover.com": "Financial",
    "americanexpress.com": "Financial",
    "ally.com": "Financial",
    # Bills
    "spectrum.net": "Bills",
    "att.com": "Bills",
    "verizon.com": "Bills",
    "tmobile.com": "Bills",
    "xfinity.com": "Bills",
    "comcast.com": "Bills",
    "duke-energy.com": "Bills",
    "aep.com": "Bills",
    "firstenergycorp.com": "Bills",
    "geico.com": "Bills",
    "statefarm.com": "Bills",
    "progressive.com": "Bills",
    "allstate.com": "Bills",
    # Promotional
    "target.com": "Promotional",
    "walmart.com": "Promotional",
    "amazon.com": "Promotional",
    "bestbuy.com": "Promotional",
    "homedepot.com": "Promotional",
    "lowes.com": "Promotional",
    "kroger.com": "Promotional",
    "krogermail.com": "Promotional",
    "costco.com": "Promotional",
    "kohls.com": "Promotional",
    "macys.com": "Promotional",
    "nike.com": "Promotional",
    "underarmour.com": "Promotional",
    "adidas.com": "Promotional",
    "delta.com": "Promotional",
    "southwest.com": "Promotional",
    "united.com": "Promotional",
    "dairyqueen.com": "Promotional",
    "linkedin.com": "Promotional",
    "facebook.com": "Promotional",
    "facebookmail.com": "Promotional",
    "twitter.com": "Promotional",
    "x.com": "Promotional",
    "instagram.com": "Promotional",
    "tiktok.com": "Promotional",
    "youtube.com": "Promotional",
    "spotify.com": "Promotional",
    "apple.com": "Promotional",
    "medallia.com": "Promotional",
    "rockler.com": "Promotional",
    "express.medallia.com": "Promotional",
    "beehiiv.com": "Promotional",
    "medium.com": "Promotional",
    "substack.com": "Promotional",
    "footballguys.com": "Promotional",
    "tldrnewsletter.com": "Promotional",
    "disneyplus.com": "Promotional",
    "netflix.com": "Promotional",
    "hulu.com": "Promotional",
    "ballotpedia.org": "Promotional",
}

# Subject keyword patterns → folder. Checked after sender rules.
_SUBJECT_PATTERNS: list[tuple[str, str]] = [
    # Financial
    (r"bank\s*statement|investment|tax\s*(document|return|form)|payment\s*confirm|transfer|credit\s*(card|score)|trading", "Financial"),
    # Bills
    (r"(bill|invoice|payment)\s*(due|reminder|notice)|subscription\s*renew|service\s*charge|auto.?pay|statement\s*ready", "Bills"),
    # Promotional
    (r"unsubscribe|% off|\bsale\b|coupon|promo|deal|newsletter|limited.?time|order\s*(confirm|ship)|delivery\s*update|survey|reward|loyalty", "Promotional"),
]

import re as _re

def _classify_email(from_address: str, subject: str) -> tuple[str, bool]:
    """Classify an email into a folder based on sender domain and subject patterns.
    Checks dynamic server-side rules first (cached 5 min), then the hardcoded
    fallback dict, then subject patterns.
    Returns (folder_name, matched) — matched=False means no rule hit and
    the email needs model review instead of auto-sorting."""
    addr = (from_address or "").lower().strip()

    if "@" in addr:
        domain = addr.split("@", 1)[1]
        parts = domain.split(".")
        parent = ".".join(parts[-2:]) if len(parts) > 2 else domain

        # Check dynamic server-side rules first
        dynamic = _get_dynamic_sender_rules()
        if domain in dynamic:
            return dynamic[domain], True
        if parent != domain and parent in dynamic:
            return dynamic[parent], True

        # Fall back to hardcoded defaults
        if domain in _SENDER_RULES_FALLBACK:
            return _SENDER_RULES_FALLBACK[domain], True
        if parent != domain and parent in _SENDER_RULES_FALLBACK:
            return _SENDER_RULES_FALLBACK[parent], True

    # Check subject patterns
    subj_lower = (subject or "").lower()
    for pattern, folder in _SUBJECT_PATTERNS:
        if _re.search(pattern, subj_lower):
            return folder, True

    return "Personal", False


def _auto_organize(batch_size: int = 100, dry_run: bool = False) -> dict:
    """Fetch emails from inbox, classify by rules, and move matched ones in bulk.

    Two-pass design:
    - Pass 1 (this tool): deterministic rules sort obvious emails by sender
      domain and subject keywords. Moved immediately, no model needed.
    - Pass 2 (model): unmatched emails are returned in the response so the
      model can review and sort the ambiguous ones via outlook_bulk_move.

    Returns a summary of what was moved plus a compact list of unmatched
    emails for model review."""

    # Ensure target folders exist
    folders = _list_folders()
    folder_names = {f["name"].lower(): f["id"] for f in folders}
    needed = ["Personal", "Financial", "Bills", "Promotional"]
    for name in needed:
        if name.lower() not in folder_names:
            result = _create_folder(name)
            folder_names[name.lower()] = result["id"]

    # Fetch batch from inbox (compact — we only need sender and subject)
    emails = _list_messages(folder="inbox", max_results=batch_size, compact=True)

    if not emails:
        return {"total": 0, "message": "Inbox is empty or no emails to process."}

    # Classify each email — separate matched (auto-sort) from unmatched (model review)
    buckets: dict[str, list[str]] = {}
    matched_list: list[dict] = []
    unmatched_list: list[dict] = []
    for email in emails:
        folder, matched = _classify_email(email["from_address"], email["subject"])
        if matched:
            buckets.setdefault(folder, []).append(email["message_id"])
            matched_list.append({
                "subject": email["subject"][:60],
                "from": email["from_address"],
                "folder": folder,
            })
        else:
            unmatched_list.append({
                "message_id": email["message_id"],
                "subject": email["subject"][:80],
                "from": email["from"],
                "from_address": email["from_address"],
            })

    if dry_run:
        summary = {f: len(ids) for f, ids in buckets.items()}
        return {
            "total": len(emails),
            "dry_run": True,
            "auto_sorted": sum(summary.values()),
            "needs_review": len(unmatched_list),
            "summary": summary,
            "sample_matched": matched_list[:10],
            "unmatched": unmatched_list,
        }

    # Move matched emails in bulk per folder
    move_results = {}
    total_moved = 0
    total_failed = 0
    for folder, msg_ids in buckets.items():
        result = _bulk_move_messages(msg_ids, folder)
        move_results[folder] = result
        total_moved += result["moved"]
        total_failed += result["failed"]

    remaining = _graph_get("/me/mailFolders/inbox", {"$select": "totalItemCount"})
    inbox_remaining = remaining.get("totalItemCount", "?")

    return {
        "total_processed": len(emails),
        "auto_moved": total_moved,
        "total_failed": total_failed,
        "by_folder": {f: r["moved"] for f, r in move_results.items()},
        "errors": [e for r in move_results.values() for e in r.get("errors", [])],
        "needs_review": len(unmatched_list),
        "unmatched": unmatched_list,
        "inbox_remaining": inbox_remaining,
        "has_more": inbox_remaining != "?" and int(inbox_remaining) > 0,
    }


# ── Inbox rules ────────────────────────────────────────────────────────────

def _list_rules() -> list:
    """List inbox message rules."""
    data = _graph_get("/me/mailFolders/inbox/messageRules")
    results = []
    for r in data.get("value", []):
        results.append({
            "id": r["id"],
            "name": r.get("displayName", ""),
            "sequence": r.get("sequence", 0),
            "is_enabled": r.get("isEnabled", False),
            "conditions": r.get("conditions", {}),
            "actions": r.get("actions", {}),
            "exceptions": r.get("exceptions", {}),
        })
    return results


def _get_rule(rule_id: str) -> dict:
    """Get a single rule by ID."""
    r = _graph_get(f"/me/mailFolders/inbox/messageRules/{rule_id}")
    return {
        "id": r["id"],
        "name": r.get("displayName", ""),
        "sequence": r.get("sequence", 0),
        "is_enabled": r.get("isEnabled", False),
        "is_read_only": r.get("isReadOnly", False),
        "conditions": r.get("conditions", {}),
        "actions": r.get("actions", {}),
        "exceptions": r.get("exceptions", {}),
    }


def _validate_rule_conditions(conditions: dict, existing_rules: list) -> list[str]:
    """Check for common rule problems. Returns list of warning strings."""
    warnings = []

    # Check senderContains for whitespace
    for val in conditions.get("senderContains", []):
        stripped = val.strip()
        if stripped != val:
            warnings.append(
                f"senderContains '{val}' has leading/trailing whitespace"
            )
        if " " in stripped:
            warnings.append(
                f"senderContains '{stripped}' contains spaces — likely malformed"
            )

    # Check headerContains for overly broad values
    for val in conditions.get("headerContains", []):
        if len(val) <= 3:
            warnings.append(
                f"headerContains '{val}' is very short — may match unintended emails"
            )

    # Check for domain overlap with existing senderContains rules
    new_domains = {s.upper().strip() for s in conditions.get("senderContains", [])}
    if new_domains:
        for rule in existing_rules:
            existing_domains = {
                s.upper().strip()
                for s in rule.get("conditions", {}).get("senderContains", [])
            }
            overlap = new_domains & existing_domains
            if overlap:
                warnings.append(
                    f"senderContains overlap with rule '{rule['name']}': "
                    f"{', '.join(overlap)}"
                )

    # Warn if headerContains may be shadowed by an existing senderContains
    if conditions.get("headerContains"):
        for rule in existing_rules:
            rule_domains = rule.get("conditions", {}).get("senderContains", [])
            if rule_domains and rule.get("actions", {}).get("stopProcessingRules"):
                warnings.append(
                    f"headerContains rule may be shadowed by senderContains "
                    f"in rule '{rule['name']}' (seq {rule.get('sequence', '?')}). "
                    f"Ensure the domain is not in that rule's senderContains."
                )
                break  # One warning is enough

    return warnings


def _create_rule(display_name: str, conditions: dict, actions: dict,
                 exceptions: dict = None, is_enabled: bool = True) -> dict:
    """Create a new inbox rule. Auto-assigns the next sequence number.
    Validates conditions and logs warnings for potential problems."""
    # Graph API requires a non-zero sequence number
    existing = _list_rules()

    # Validate and log warnings
    warnings = _validate_rule_conditions(conditions, existing)
    for w in warnings:
        logger.warning(f"Rule '{display_name}': {w}")

    next_seq = max((r.get("sequence", 0) for r in existing), default=0) + 1
    body = {
        "displayName": display_name,
        "sequence": next_seq,
        "isEnabled": is_enabled,
        "conditions": conditions,
        "actions": actions,
    }
    if exceptions:
        body["exceptions"] = exceptions
    data = _graph_post("/me/mailFolders/inbox/messageRules", body)
    result = {
        "id": data.get("id", ""),
        "name": data.get("displayName", display_name),
    }
    if warnings:
        result["warnings"] = warnings
    _invalidate_rules_cache()
    return result


def _update_rule(rule_id: str, updates: dict) -> dict:
    """Update an existing rule. Pass any subset of displayName, conditions, actions, exceptions, isEnabled."""
    result = _graph_patch(f"/me/mailFolders/inbox/messageRules/{rule_id}", updates)
    _invalidate_rules_cache()
    return result


def _delete_rule(rule_id: str) -> bool:
    """Delete an inbox rule."""
    result = _graph_delete(f"/me/mailFolders/inbox/messageRules/{rule_id}")
    _invalidate_rules_cache()
    return result


# ── MCP Tool Registration ─────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        Tool(
            name="outlook_list_emails",
            description=(
                "List emails from your Outlook / Hotmail inbox (or another folder). "
                "Returns subject, sender, date, read status, and a short preview. "
                "Use compact=true for bulk sorting — returns only id, subject, and sender "
                "(much smaller output, fits more emails per batch)."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {
                        "type": "string",
                        "description": "Mail folder to list (default: inbox). Use outlook_list_folders to see available folders.",
                        "default": "inbox",
                    },
                    "max_results": {
                        "type": "integer",
                        "description": "Maximum number of emails to return (default: 20, max: 100)",
                        "default": 20,
                    },
                    "unread_only": {
                        "type": "boolean",
                        "description": "Only show unread emails (default: false)",
                        "default": False,
                    },
                    "compact": {
                        "type": "boolean",
                        "description": "Return minimal data (id, subject, sender only). Use for bulk sorting to reduce context size.",
                        "default": False,
                    },
                },
                "required": [],
            },
        ),
        Tool(
            name="outlook_read_email",
            description=(
                "Read the full content of a specific Outlook email by its message ID "
                "(from outlook_list_emails). Returns subject, sender, recipients, date, "
                "and the full body text."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {
                        "type": "string",
                        "description": "Message ID from outlook_list_emails results",
                    },
                },
                "required": ["message_id"],
            },
        ),
        Tool(
            name="outlook_search_emails",
            description=(
                "Search Outlook emails by keyword. Searches across all folders — subject, "
                "body, sender, etc. Use when looking for a specific person, topic, or thread."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "Search query (e.g., sender name, subject keyword, topic)",
                    },
                    "max_results": {
                        "type": "integer",
                        "description": "Maximum results (default: 20)",
                        "default": 20,
                    },
                },
                "required": ["query"],
            },
        ),
        Tool(
            name="outlook_send_email",
            description=(
                "Send a new email from your Outlook / Hotmail account. "
                "For replying to an existing thread, use outlook_reply instead."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "string", "description": "Recipient email address(es), comma-separated"},
                    "subject": {"type": "string", "description": "Email subject line"},
                    "body": {"type": "string", "description": "Plain text body"},
                    "cc": {"type": "string", "description": "CC address(es), comma-separated (optional)"},
                },
                "required": ["to", "subject", "body"],
            },
        ),
        Tool(
            name="outlook_reply",
            description=(
                "Reply to an existing Outlook email by message ID. Automatically threads "
                "the reply. Set reply_all=true to reply to all recipients."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {
                        "type": "string",
                        "description": "Message ID of the email to reply to",
                    },
                    "body": {"type": "string", "description": "Reply body text"},
                    "reply_all": {
                        "type": "boolean",
                        "description": "Reply to all recipients (default: false)",
                        "default": False,
                    },
                },
                "required": ["message_id", "body"],
            },
        ),
        Tool(
            name="outlook_mark_read",
            description="Mark an Outlook email as read or unread.",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                    "read": {
                        "type": "boolean",
                        "description": "True to mark read, false for unread (default: true)",
                        "default": True,
                    },
                },
                "required": ["message_id"],
            },
        ),
        Tool(
            name="outlook_move_email",
            description=(
                "Move an Outlook email to a different folder. Use outlook_list_folders "
                "to see available folder names."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                    "folder": {"type": "string", "description": "Destination folder name (e.g., Archive, Junk Email)"},
                },
                "required": ["message_id", "folder"],
            },
        ),
        Tool(
            name="outlook_delete_email",
            description="Delete an Outlook email. Moves to Deleted Items by default; pass permanent=true to hard-delete.",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                    "permanent": {
                        "type": "boolean",
                        "description": "Hard-delete instead of moving to Deleted Items",
                        "default": False,
                    },
                },
                "required": ["message_id"],
            },
        ),
        Tool(
            name="outlook_list_folders",
            description="List all mail folders in the Outlook account with unread and total counts.",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": [],
            },
        ),
        Tool(
            name="outlook_list_attachments",
            description="List attachments on a specific email. Returns attachment names, types, and sizes.",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                },
                "required": ["message_id"],
            },
        ),
        # ── Categories ──
        Tool(
            name="outlook_list_categories",
            description="List all available Outlook categories (colored labels) for this account.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="outlook_categorize_email",
            description=(
                "Assign or remove category labels on an email. Pass a list of category names "
                "to set, or an empty list to clear all categories."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                    "categories": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of category names to assign (e.g. ['Red category', 'Blue category']). Empty list to clear.",
                    },
                },
                "required": ["message_id", "categories"],
            },
        ),
        # ── Folder management ──
        Tool(
            name="outlook_create_folder",
            description="Create a new mail folder. Optionally provide a parent folder name/ID to create a subfolder.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Name for the new folder"},
                    "parent_folder": {"type": "string", "description": "Parent folder name or ID for creating a subfolder (optional)"},
                },
                "required": ["name"],
            },
        ),
        Tool(
            name="outlook_rename_folder",
            description="Rename an existing mail folder. Use outlook_list_folders to find folder IDs.",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {"type": "string", "description": "Folder name or ID to rename"},
                    "new_name": {"type": "string", "description": "New name for the folder"},
                },
                "required": ["folder", "new_name"],
            },
        ),
        Tool(
            name="outlook_delete_folder",
            description="Delete a mail folder. Use outlook_list_folders to find folder IDs.",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {"type": "string", "description": "Folder name or ID to delete"},
                },
                "required": ["folder"],
            },
        ),
        Tool(
            name="outlook_bulk_move",
            description=(
                "Move multiple emails to a folder in one call. Much faster than calling "
                "outlook_move_email repeatedly. Pass an array of message IDs and the "
                "destination folder name. Returns counts of successful and failed moves."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "message_ids": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of message IDs to move",
                    },
                    "folder": {
                        "type": "string",
                        "description": "Destination folder name (e.g. Promotional, Financial, Bills, Personal)",
                    },
                },
                "required": ["message_ids", "folder"],
            },
        ),
        # ── Inbox rules ──
        Tool(
            name="outlook_list_rules",
            description="List all inbox rules. Shows each rule's name, conditions, actions, and enabled status.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        Tool(
            name="outlook_get_rule",
            description="Get full details of a specific inbox rule by ID.",
            inputSchema={
                "type": "object",
                "properties": {
                    "rule_id": {"type": "string", "description": "Rule ID from outlook_list_rules"},
                },
                "required": ["rule_id"],
            },
        ),
        Tool(
            name="outlook_create_rule",
            description=(
                "Create a new inbox rule. Conditions and actions use Microsoft Graph messageRule format. "
                "Common conditions: senderContains, subjectContains, fromAddresses, hasAttachments, importance. "
                "Common actions: moveToFolder, copyToFolder, delete, markAsRead, stopProcessingRules, "
                "forwardTo, categories."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Display name for the rule"},
                    "conditions": {
                        "type": "object",
                        "description": "Rule conditions (e.g. {\"senderContains\": [\"example.com\"]})",
                    },
                    "actions": {
                        "type": "object",
                        "description": "Rule actions (e.g. {\"moveToFolder\": \"folder_id\", \"stopProcessingRules\": true})",
                    },
                    "exceptions": {
                        "type": "object",
                        "description": "Exception conditions — messages matching these are excluded (optional)",
                    },
                    "is_enabled": {
                        "type": "boolean",
                        "description": "Whether the rule is active (default: true)",
                        "default": True,
                    },
                },
                "required": ["name", "conditions", "actions"],
            },
        ),
        Tool(
            name="outlook_update_rule",
            description=(
                "Update an existing inbox rule. Pass any fields to change: "
                "displayName, conditions, actions, exceptions, isEnabled."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "rule_id": {"type": "string", "description": "Rule ID from outlook_list_rules"},
                    "updates": {
                        "type": "object",
                        "description": "Fields to update (e.g. {\"isEnabled\": false} or {\"actions\": {\"markAsRead\": true}})",
                    },
                },
                "required": ["rule_id", "updates"],
            },
        ),
        Tool(
            name="outlook_delete_rule",
            description="Delete an inbox rule by ID.",
            inputSchema={
                "type": "object",
                "properties": {
                    "rule_id": {"type": "string", "description": "Rule ID from outlook_list_rules"},
                },
                "required": ["rule_id"],
            },
        ),
        # ── Auto-organize ──
        Tool(
            name="outlook_auto_organize",
            description=(
                "Automatically sort inbox emails into folders (Promotional, Financial, Bills, Personal) "
                "using built-in rules based on sender domain and subject keywords. No manual sorting needed — "
                "the server handles everything. Creates missing folders automatically. "
                "Use dry_run=true to preview what would be moved without actually moving. "
                "Call repeatedly until has_more is false to process the entire inbox."
            ),
            inputSchema={
                "type": "object",
                "properties": {
                    "batch_size": {
                        "type": "integer",
                        "description": "Number of emails to process per batch (default: 100, max: 100)",
                        "default": 100,
                    },
                    "dry_run": {
                        "type": "boolean",
                        "description": "Preview mode — show what would be moved without actually moving (default: false)",
                        "default": False,
                    },
                },
                "required": [],
            },
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    try:
        if name == "outlook_list_emails":
            folder = arguments.get("folder", "inbox")
            max_results = arguments.get("max_results", 20)
            unread_only = arguments.get("unread_only", False)
            compact = arguments.get("compact", False)
            results = _list_messages(folder, max_results, unread_only, compact=compact)
            if not results:
                return [TextContent(type="text", text="No emails found.")]
            lines = [f"Found {len(results)} email(s):\n"]
            for i, em in enumerate(results, 1):
                if compact:
                    lines.append(
                        f"{i}. {em['subject']} | {em['from']} ({em['from_address']}) | {em['message_id']}"
                    )
                else:
                    read_mark = "" if em.get("is_read", True) else " [UNREAD]"
                    attach = " 📎" if em.get("has_attachments") else ""
                    lines.append(
                        f"{i}. **{em['subject']}**{read_mark}{attach}\n"
                        f"   From: {em['from']} ({em['from_address']})\n"
                        f"   Date: {em.get('date', '')}\n"
                        f"   message_id: {em['message_id']}\n"
                        f"   Preview: {em.get('preview', '')}"
                    )
            return [TextContent(type="text", text="\n".join(lines) if compact else "\n\n".join(lines))]

        elif name == "outlook_read_email":
            msg_id = arguments.get("message_id")
            if not msg_id:
                return [TextContent(type="text", text="Error: message_id is required")]
            result = _read_message(msg_id)
            text = (
                f"**Subject:** {result['subject']}\n"
                f"**From:** {result['from']} ({result['from_address']})\n"
                f"**To:** {result['to']}\n"
            )
            if result['cc']:
                text += f"**Cc:** {result['cc']}\n"
            text += (
                f"**Date:** {result['date']}\n"
                f"**message_id:** {result['message_id']}\n"
                f"**Read:** {result['is_read']}\n"
                f"**Has Attachments:** {result['has_attachments']}\n"
            )
            if result['has_attachments']:
                text += "\n_Use `outlook_list_attachments` with this message ID to see attachments._\n"
            text += f"\n---\n\n{result['body']}"
            return [TextContent(type="text", text=text)]

        elif name == "outlook_search_emails":
            query = arguments.get("query", "")
            max_results = arguments.get("max_results", 20)
            if not query:
                return [TextContent(type="text", text="Error: query is required")]
            results = _search_messages(query, max_results)
            if not results:
                return [TextContent(type="text", text=f'No emails matched "{query}".')]
            lines = [f'Found {len(results)} email(s) matching "{query}":\n']
            for i, em in enumerate(results, 1):
                read_mark = "" if em["is_read"] else " [UNREAD]"
                lines.append(
                    f"{i}. **{em['subject']}**{read_mark}\n"
                    f"   From: {em['from']} ({em['from_address']})\n"
                    f"   Date: {em['date']}\n"
                    f"   message_id: {em['message_id']}\n"
                    f"   Preview: {em['preview']}"
                )
            return [TextContent(type="text", text="\n\n".join(lines))]

        elif name == "outlook_send_email":
            to = arguments.get("to")
            subject = arguments.get("subject")
            body = arguments.get("body")
            if not to or not subject or body is None:
                return [TextContent(type="text", text="Error: to, subject, and body are required")]
            _send_message(to, subject, body, cc=arguments.get("cc"))
            return [TextContent(type="text", text=f"Email sent to {to} with subject '{subject}'.")]

        elif name == "outlook_reply":
            msg_id = arguments.get("message_id")
            body = arguments.get("body")
            if not msg_id or body is None:
                return [TextContent(type="text", text="Error: message_id and body are required")]
            reply_all = bool(arguments.get("reply_all", False))
            _reply_to_message(msg_id, body, reply_all)
            mode = "Reply-all" if reply_all else "Reply"
            return [TextContent(type="text", text=f"{mode} sent for message {msg_id}.")]

        elif name == "outlook_mark_read":
            msg_id = arguments.get("message_id")
            if not msg_id:
                return [TextContent(type="text", text="Error: message_id is required")]
            read = bool(arguments.get("read", True))
            _mark_read(msg_id, read)
            state = "read" if read else "unread"
            return [TextContent(type="text", text=f"Marked message as {state}.")]

        elif name == "outlook_move_email":
            msg_id = arguments.get("message_id")
            folder = arguments.get("folder")
            if not msg_id or not folder:
                return [TextContent(type="text", text="Error: message_id and folder are required")]
            _move_message(msg_id, folder)
            return [TextContent(type="text", text=f"Moved message to {folder}.")]

        elif name == "outlook_delete_email":
            msg_id = arguments.get("message_id")
            if not msg_id:
                return [TextContent(type="text", text="Error: message_id is required")]
            permanent = bool(arguments.get("permanent", False))
            ok = _delete_message(msg_id, permanent)
            action = "Permanently deleted" if permanent else "Moved to Deleted Items"
            return [TextContent(type="text", text=f"{action}: {msg_id}")]

        elif name == "outlook_list_folders":
            folders = _list_folders()
            if not folders:
                return [TextContent(type="text", text="No folders found.")]
            lines = [f"Found {len(folders)} folder(s):\n"]
            for f in folders:
                unread = f"  ({f['unread_count']} unread)" if f['unread_count'] else ""
                lines.append(f"- **{f['name']}** — {f['total_count']} total{unread}\n  ID: {f['id']}")
            return [TextContent(type="text", text="\n".join(lines))]

        elif name == "outlook_list_attachments":
            msg_id = arguments.get("message_id")
            if not msg_id:
                return [TextContent(type="text", text="Error: message_id is required")]
            attachments = _list_attachments(msg_id)
            if not attachments:
                return [TextContent(type="text", text="No attachments found on this message.")]
            lines = [f"Found {len(attachments)} attachment(s):\n"]
            for a in attachments:
                size_kb = a['size'] // 1024
                lines.append(f"- **{a['name']}** ({a['content_type']}, {size_kb}KB)\n  ID: {a['id']}")
            return [TextContent(type="text", text="\n".join(lines))]

        # ── Categories ──
        elif name == "outlook_list_categories":
            cats = _list_categories()
            if not cats:
                return [TextContent(type="text", text="No categories found.")]
            lines = [f"Found {len(cats)} category/categories:\n"]
            for c in cats:
                lines.append(f"- **{c['name']}** (color: {c['color']})")
            return [TextContent(type="text", text="\n".join(lines))]

        elif name == "outlook_categorize_email":
            msg_id = arguments.get("message_id")
            categories = arguments.get("categories", [])
            if not msg_id:
                return [TextContent(type="text", text="Error: message_id is required")]
            _categorize_message(msg_id, categories)
            if categories:
                return [TextContent(type="text", text=f"Set categories: {', '.join(categories)}")]
            return [TextContent(type="text", text="Cleared all categories.")]

        # ── Folder management ──
        elif name == "outlook_create_folder":
            folder_name = arguments.get("name")
            if not folder_name:
                return [TextContent(type="text", text="Error: name is required")]
            parent = arguments.get("parent_folder")
            parent_id = _resolve_folder_id(parent) if parent else None
            result = _create_folder(folder_name, parent_id)
            return [TextContent(type="text", text=f"Created folder **{result['name']}** (ID: {result['id']})")]

        elif name == "outlook_rename_folder":
            folder = arguments.get("folder")
            new_name = arguments.get("new_name")
            if not folder or not new_name:
                return [TextContent(type="text", text="Error: folder and new_name are required")]
            folder_id = _resolve_folder_id(folder)
            _rename_folder(folder_id, new_name)
            return [TextContent(type="text", text=f"Renamed folder to **{new_name}**")]

        elif name == "outlook_delete_folder":
            folder = arguments.get("folder")
            if not folder:
                return [TextContent(type="text", text="Error: folder is required")]
            folder_id = _resolve_folder_id(folder)
            _delete_folder(folder_id)
            return [TextContent(type="text", text=f"Deleted folder.")]

        elif name == "outlook_bulk_move":
            msg_ids = arguments.get("message_ids", [])
            folder = arguments.get("folder")
            if not msg_ids or not folder:
                return [TextContent(type="text", text="Error: message_ids and folder are required")]
            result = _bulk_move_messages(msg_ids, folder)
            text = f"Moved {result['moved']} email(s) to **{folder}**."
            if result['failed']:
                text += f" Failed: {result['failed']}."
                if result['errors']:
                    text += "\n" + "\n".join(result['errors'][:3])
            return [TextContent(type="text", text=text)]

        # ── Inbox rules ──
        elif name == "outlook_list_rules":
            rules = _list_rules()
            if not rules:
                return [TextContent(type="text", text="No inbox rules found.")]
            lines = [f"Found {len(rules)} rule(s):\n"]
            for r in rules:
                enabled = "enabled" if r["is_enabled"] else "disabled"
                lines.append(
                    f"**{r['name']}** ({enabled}, seq {r['sequence']})\n"
                    f"  ID: {r['id']}\n"
                    f"  Conditions: {json.dumps(r['conditions'], indent=None)}\n"
                    f"  Actions: {json.dumps(r['actions'], indent=None)}"
                )
            return [TextContent(type="text", text="\n\n".join(lines))]

        elif name == "outlook_get_rule":
            rule_id = arguments.get("rule_id")
            if not rule_id:
                return [TextContent(type="text", text="Error: rule_id is required")]
            r = _get_rule(rule_id)
            text = (
                f"**{r['name']}**\n"
                f"ID: {r['id']}\n"
                f"Enabled: {r['is_enabled']}\n"
                f"Read-only: {r['is_read_only']}\n"
                f"Sequence: {r['sequence']}\n"
                f"Conditions: {json.dumps(r['conditions'], indent=2)}\n"
                f"Actions: {json.dumps(r['actions'], indent=2)}\n"
                f"Exceptions: {json.dumps(r['exceptions'], indent=2)}"
            )
            return [TextContent(type="text", text=text)]

        elif name == "outlook_create_rule":
            rule_name = arguments.get("name")
            conditions = arguments.get("conditions", {})
            actions = arguments.get("actions", {})
            exceptions = arguments.get("exceptions")
            is_enabled = arguments.get("is_enabled", True)
            if not rule_name:
                return [TextContent(type="text", text="Error: name is required")]
            result = _create_rule(rule_name, conditions, actions, exceptions, is_enabled)
            text = f"Created rule **{result['name']}** (ID: {result['id']})"
            if result.get("warnings"):
                text += "\n\n⚠️ Warnings:\n" + "\n".join(
                    f"- {w}" for w in result["warnings"]
                )
            return [TextContent(type="text", text=text)]

        elif name == "outlook_update_rule":
            rule_id = arguments.get("rule_id")
            updates = arguments.get("updates", {})
            if not rule_id:
                return [TextContent(type="text", text="Error: rule_id is required")]
            _update_rule(rule_id, updates)
            return [TextContent(type="text", text=f"Updated rule {rule_id}.")]

        elif name == "outlook_delete_rule":
            rule_id = arguments.get("rule_id")
            if not rule_id:
                return [TextContent(type="text", text="Error: rule_id is required")]
            _delete_rule(rule_id)
            return [TextContent(type="text", text=f"Deleted rule {rule_id}.")]

        # ── Auto-organize ──
        elif name == "outlook_auto_organize":
            batch_size = min(arguments.get("batch_size", 100), 100)
            dry_run = arguments.get("dry_run", False)
            result = _auto_organize(batch_size=batch_size, dry_run=dry_run)

            if result.get("total", 0) == 0:
                return [TextContent(type="text", text=result.get("message", "No emails to process."))]

            if dry_run:
                lines = [f"**DRY RUN** — {result['total']} email(s) scanned:\n"]
                lines.append(f"Auto-sortable: {result['auto_sorted']}")
                lines.append(f"Needs model review: {result['needs_review']}\n")
                lines.append("**Auto-sort breakdown:**")
                for folder, count in sorted(result.get("summary", {}).items()):
                    lines.append(f"  {folder}: {count}")
            else:
                lines = [f"**Auto-sorted {result['auto_moved']}** of {result['total_processed']} email(s):\n"]
                for folder, count in sorted(result.get("by_folder", {}).items()):
                    lines.append(f"  {folder}: {count} moved")
                if result.get("total_failed"):
                    lines.append(f"\nFailed: {result['total_failed']}")

            # Append unmatched emails for model review
            unmatched = result.get("unmatched", [])
            if unmatched:
                lines.append(f"\n**{len(unmatched)} email(s) need your review** — no rule matched:")
                lines.append("message_id | subject | from")
                lines.append("--- | --- | ---")
                for u in unmatched:
                    lines.append(f"{u['message_id']} | {u['subject']} | {u['from']} ({u['from_address']})")
                lines.append(
                    "\nSort these into the correct folders using outlook_bulk_move. "
                    "If you see a pattern (e.g., a sender that should always go to Promotional), "
                    "tell the user so they can add it to the server's rules."
                )
            else:
                lines.append("\nAll emails matched rules — no review needed.")

            if not dry_run:
                lines.append(f"\nInbox remaining: {result['inbox_remaining']}")
                if result.get("has_more"):
                    lines.append("Call again to process the next batch.")

            return [TextContent(type="text", text="\n".join(lines))]

        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        logger.exception("Tool %s failed", name)
        return [TextContent(type="text", text=f"Error: {e}")]


# ── Main ──

async def run():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream, write_stream, server.create_initialization_options()
        )


if __name__ == "__main__":
    # Authenticate before entering the MCP stdio loop so the browser
    # opens immediately on first launch rather than on the first tool call.
    try:
        _acquire_token()
    except Exception as e:
        logger.error("Startup authentication failed: %s", e)
        sys.exit(1)

    asyncio.run(run())
