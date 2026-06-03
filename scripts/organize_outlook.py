#!/usr/bin/env python3
"""
organize_outlook.py — Two-model email categorization pipeline.

Pipeline stages:
  1. FETCH:  Pull emails from inbox via Graph API → save raw metadata to JSON
  2. CATEGORIZE (Model 1):  LLM reads subject + sender + attachment names,
     assigns a free-form category to each email. No constraints — the model
     discovers whatever categories make sense.
  3. CONSOLIDATE (Model 2):  Second LLM reviews the raw category list and
     proposes a practical folder structure — merging similar labels, naming
     top-level folders, mapping each category into the structure.
  4. MOVE:  Python applies the folder map — creates folders as needed,
     bulk-moves emails by message_id. No model involved.

Usage:
  python scripts/organize_outlook.py                    # full pipeline
  python scripts/organize_outlook.py --fetch-only       # just fetch, save JSON
  python scripts/organize_outlook.py --dry-run          # everything except moves
  python scripts/organize_outlook.py --batch-size 50    # smaller batches
  python scripts/organize_outlook.py --resume           # resume from saved JSON
  python scripts/organize_outlook.py --model1 qwen3:30b-a3b --model2 gemma4:12b

Requires:
  - OUTLOOK_CLIENT_ID env var (or cached token)
  - Ollama running locally (or set --endpoint for remote)
"""

import argparse
import json
import logging
import os
import sys
import time
from pathlib import Path

# Add project root to path so we can import the outlook server module
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import io as _io
if hasattr(sys.stdout, "buffer"):
    sys.stdout = _io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "buffer"):
    sys.stderr = _io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# ── Config ────────────────────────────────────────────────────────────────

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
PIPELINE_DIR = DATA_DIR / "outlook_pipeline"

DEFAULT_ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"
DEFAULT_MODEL1 = "groq/compound"      # categorizer — bigger model, handles 100 emails fully
DEFAULT_MODEL2 = "groq/compound-mini" # consolidator — small input (category list only)


# ── Stage 1: Fetch ────────────────────────────────────────────────────────

def fetch_emails(batch_size: int = 100) -> list[dict]:
    """Fetch emails from inbox with subject, sender, and attachment names.
    No body content — keeps tokens low and avoids injection surface."""
    from mcp_servers.outlook_server import _list_messages

    logger.info(f"Fetching up to {batch_size} emails from inbox...")
    emails = _list_messages(
        folder="inbox",
        max_results=batch_size,
        compact=True,
        include_attachment_names=True,
    )
    logger.info(f"Fetched {len(emails)} emails")
    return emails


def save_stage(filename: str, data: any) -> Path:
    """Save pipeline stage output to JSON."""
    PIPELINE_DIR.mkdir(parents=True, exist_ok=True)
    path = PIPELINE_DIR / filename
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    logger.info(f"Saved {path} ({len(json.dumps(data))} bytes)")
    return path


def load_stage(filename: str) -> any:
    """Load pipeline stage output from JSON."""
    path = PIPELINE_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"Stage file not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


# ── LLM Call Helper ───────────────────────────────────────────────────────

def llm_call(endpoint: str, model: str, system: str, user: str,
             temperature: float = 0.3, max_tokens: int = 8192,
             api_key: str = None) -> str:
    """Call an OpenAI-compatible chat endpoint. Returns the assistant text."""
    import httpx

    messages = [
        {"role": "system", "content": system},
        {"role": "user", "content": user},
    ]
    body = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
    }
    headers = {}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    logger.info(f"Calling {model} ({len(user)} chars input)...")
    t0 = time.time()
    with httpx.Client(timeout=600) as client:
        for attempt in range(5):
            resp = client.post(endpoint, json=body, headers=headers)
            if resp.status_code == 429:
                # Parse Retry-After header if present, else exponential backoff
                wait = int(resp.headers.get("retry-after", 2 ** (attempt + 2)))
                logger.warning(f"Rate limited — waiting {wait}s before retry (attempt {attempt + 1}/5)...")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            break
        else:
            resp.raise_for_status()
    result = resp.json()
    content = result["choices"][0]["message"]["content"]
    # Strip <think> blocks if present (Qwen3, DeepSeek)
    import re
    content = re.sub(r"<think>.*?</think>", "", content, flags=re.DOTALL).strip()
    elapsed = time.time() - t0
    logger.info(f"Model responded in {elapsed:.1f}s ({len(content)} chars)")
    return content


# ── Stage 2: Categorize (Model 1) ────────────────────────────────────────

CATEGORIZE_SYSTEM = """\
You are an email categorization assistant. You will receive a numbered list of \
emails with their subject line, sender, and attachment filenames (if any).

For each email, assign a short, descriptive category label that captures what \
kind of email it is. Use your best judgment — there are no predefined categories. \
Think about what folder a well-organized person would file it into.

Examples of good category labels:
  Grocery Delivery, Credit Monitoring, Tech Newsletter, Fantasy Football, \
  Military Correspondence, Retail Promotion, Shipping Notification, \
  Subscription Renewal, Bank Statement, Insurance, Social Media, \
  Software Update, Security Alert, Job Board, Utility Bill

Rules:
- One category per email
- Keep labels short (1-3 words)
- Be specific enough to be useful (not just "Email" or "Notification")
- Emails with "receipt", "order confirmed", or "order confirmation" in the \
subject are NOT promotional — they are transactional (e.g., "Order Receipt", \
"Purchase Confirmation"). Do not categorize these as promotions, deals, or ads \
even if the sender is a retailer.
- Output ONLY valid JSON — a list of objects with "index" and "category"
- Do not include any explanation or commentary, just the JSON array"""

def build_categorize_prompt(emails: list[dict]) -> str:
    """Build the user prompt for Model 1 — numbered list of emails."""
    lines = []
    for i, em in enumerate(emails, 1):
        parts = [f"{i}. Subject: {em['subject'][:80]}"]
        parts.append(f"   From: {em.get('from', '')} ({em.get('from_address', '')})")
        att_names = em.get("attachment_names", [])
        if att_names:
            parts.append(f"   Attachments: {', '.join(att_names[:5])}")
        lines.append("\n".join(parts))
    return "\n\n".join(lines)


def parse_categories(response: str, count: int) -> list[dict]:
    """Parse Model 1's JSON response into a list of {index, category}."""
    # Find the JSON array in the response (DOTALL so it spans newlines)
    import re
    match = re.search(r"\[.*\]", response, re.DOTALL)
    if not match:
        # Response might be truncated (no closing ]). Try to salvage by
        # finding the start and appending a closing bracket.
        match = re.search(r"\[.*", response, re.DOTALL)
        if not match:
            raise ValueError(f"No JSON array found in model response:\n{response[:500]}")
        # Trim to the last complete object and close the array
        raw = match.group().rstrip().rstrip(",")
        # Find the last complete } and truncate there
        last_brace = raw.rfind("}")
        if last_brace == -1:
            raise ValueError(f"No complete JSON objects found in truncated response:\n{response[:500]}")
        raw = raw[:last_brace + 1] + "]"
        logger.warning(f"JSON array was truncated — salvaged partial response")
        items = json.loads(raw)
    else:
        items = json.loads(match.group())
    # Validate
    if not isinstance(items, list):
        raise ValueError(f"Expected JSON array, got {type(items)}")

    # Normalize — accept both {"index": 1, "category": "X"} and [1, "X"] formats
    result = []
    for item in items:
        if isinstance(item, dict):
            result.append({
                "index": int(item.get("index", item.get("id", 0))),
                "category": str(item.get("category", item.get("label", "Unknown"))),
            })
        elif isinstance(item, list) and len(item) >= 2:
            result.append({"index": int(item[0]), "category": str(item[1])})

    logger.info(f"Parsed {len(result)} categories from model response")
    return result


# ── Stage 3: Consolidate (Model 2) ───────────────────────────────────────

CONSOLIDATE_SYSTEM = """\
You are an email organization architect. You will receive a list of category \
labels that were assigned to emails by a first-pass categorizer, along with \
the count of emails in each category.

Your job is to design a practical folder structure:
1. Review all the categories and their counts
2. Merge similar/overlapping categories into top-level folders
3. Aim for 4-8 top-level folders — enough to be useful, few enough to navigate
4. Every original category must map to exactly one folder
5. Use clear, concise folder names

Output ONLY valid JSON with this structure:
{
  "folders": ["Folder1", "Folder2", ...],
  "reasoning": "Brief explanation of why this structure makes sense",
  "mapping": {
    "Original Category 1": "Folder1",
    "Original Category 2": "Folder1",
    "Original Category 3": "Folder2",
    ...
  }
}

Do not include any text outside the JSON object."""


def build_consolidate_prompt(categories: list[dict]) -> str:
    """Build the user prompt for Model 2 — category list with counts."""
    # Count occurrences
    counts: dict[str, int] = {}
    for item in categories:
        cat = item["category"]
        counts[cat] = counts.get(cat, 0) + 1

    lines = ["Categories assigned by the first-pass categorizer:\n"]
    for cat, count in sorted(counts.items(), key=lambda x: -x[1]):
        lines.append(f"  {cat}: {count} email(s)")
    lines.append(f"\nTotal: {sum(counts.values())} emails across {len(counts)} categories")
    return "\n".join(lines)


def parse_consolidation(response: str) -> dict:
    """Parse Model 2's JSON response into the folder structure."""
    import re
    match = re.search(r"\{.*\}", response, re.DOTALL)
    if not match:
        raise ValueError(f"No JSON object found in model response:\n{response[:500]}")

    result = json.loads(match.group())
    if "folders" not in result or "mapping" not in result:
        raise ValueError(f"Response missing 'folders' or 'mapping' keys: {list(result.keys())}")

    logger.info(
        f"Consolidated into {len(result['folders'])} folders: "
        f"{', '.join(result['folders'])}"
    )
    return result


# ── Stage 4: Move ─────────────────────────────────────────────────────────

def apply_moves(emails: list[dict], categories: list[dict],
                folder_map: dict, dry_run: bool = False) -> dict:
    """Create folders and move emails based on the consolidated mapping.
    Returns a summary of what was done."""
    from mcp_servers.outlook_server import (
        _list_folders, _create_folder, _bulk_move_messages,
    )

    # Build index→email lookup
    email_by_index = {i + 1: em for i, em in enumerate(emails)}

    # Build index→folder lookup
    category_by_index = {item["index"]: item["category"] for item in categories}
    mapping = folder_map["mapping"]

    # Group message_ids by target folder
    move_plan: dict[str, list[str]] = {}
    unmapped = []
    for idx, email in email_by_index.items():
        cat = category_by_index.get(idx)
        if not cat:
            unmapped.append(idx)
            continue
        folder = mapping.get(cat)
        if not folder:
            # Category exists but wasn't in the consolidation mapping
            folder = "Personal"
        move_plan.setdefault(folder, []).append(email["message_id"])

    # Summary for dry run or actual execution
    summary = {
        "plan": {folder: len(ids) for folder, ids in move_plan.items()},
        "total": sum(len(ids) for ids in move_plan.values()),
        "unmapped": len(unmapped),
        "folders_needed": list(move_plan.keys()),
    }

    if dry_run:
        summary["dry_run"] = True
        logger.info(f"DRY RUN — would move {summary['total']} emails into {len(move_plan)} folders")
        for folder, ids in sorted(move_plan.items()):
            logger.info(f"  {folder}: {len(ids)} emails")
        return summary

    # Create missing folders
    existing = {f["name"].lower(): f["id"] for f in _list_folders()}
    for folder_name in move_plan:
        if folder_name.lower() not in existing:
            logger.info(f"Creating folder: {folder_name}")
            result = _create_folder(folder_name)
            existing[folder_name.lower()] = result["id"]

    # Bulk move per folder
    results = {}
    total_moved = 0
    total_failed = 0
    for folder, msg_ids in move_plan.items():
        logger.info(f"Moving {len(msg_ids)} emails to {folder}...")
        result = _bulk_move_messages(msg_ids, folder)
        results[folder] = result
        total_moved += result["moved"]
        total_failed += result["failed"]

    summary["moved"] = total_moved
    summary["failed"] = total_failed
    summary["results"] = {f: r["moved"] for f, r in results.items()}
    summary["errors"] = [e for r in results.values() for e in r.get("errors", [])]

    logger.info(f"Done — moved {total_moved}, failed {total_failed}")
    return summary


# ── Stage 5: Rule Learning ────────────────────────────────────────────────

def learn_rules(emails: list[dict], categories: list[dict],
                folder_map: dict, dry_run: bool = False,
                min_occurrences: int = 2) -> dict:
    """Analyze categorized emails to discover sender domain patterns and
    create Outlook inbox rules so future emails self-sort.

    Only creates rules for domains that appeared min_occurrences+ times
    mapping to the same folder. Reads existing rules first to avoid dupes.

    Returns a summary of rules proposed/created."""
    from mcp_servers.outlook_server import (
        _list_rules, _create_rule, _list_folders, _resolve_folder_id,
    )

    # Build index→email and index→folder lookups
    email_by_index = {i + 1: em for i, em in enumerate(emails)}
    category_by_index = {item["index"]: item["category"] for item in categories}
    mapping = folder_map["mapping"]

    # Count sender domains per target folder
    domain_folder_counts: dict[tuple[str, str], int] = {}
    domain_examples: dict[str, str] = {}  # domain → example sender name
    for idx, email in email_by_index.items():
        cat = category_by_index.get(idx)
        if not cat:
            continue
        folder = mapping.get(cat, "Personal")
        addr = (email.get("from_address") or "").lower().strip()
        if "@" not in addr:
            continue
        domain = addr.split("@", 1)[1]
        # Use parent domain (e.g., kroger.com not e.krogermail.com)
        parts = domain.split(".")
        if len(parts) > 2:
            parent = ".".join(parts[-2:])
        else:
            parent = domain
        key = (parent, folder)
        domain_folder_counts[key] = domain_folder_counts.get(key, 0) + 1
        if parent not in domain_examples:
            domain_examples[parent] = email.get("from", parent)

    # Also track display names per (domain, folder) for headerContains rules
    display_name_counts: dict[tuple[str, str, str], int] = {}  # (domain, display_name, folder) -> count
    for idx, email in email_by_index.items():
        cat = category_by_index.get(idx)
        if not cat:
            continue
        folder = mapping.get(cat, "Personal")
        addr = (email.get("from_address") or "").lower().strip()
        if "@" not in addr:
            continue
        domain = addr.split("@", 1)[1]
        parts = domain.split(".")
        parent = ".".join(parts[-2:]) if len(parts) > 2 else domain
        display_name = (email.get("from") or "").strip()
        if display_name:
            key = (parent, display_name, folder)
            display_name_counts[key] = display_name_counts.get(key, 0) + 1

    # Only create rules for domains that CONSISTENTLY map to one folder.
    # If a domain splits across multiple folders (e.g., Kroger sends both
    # order confirmations and weekly ads), it's ambiguous — but we may be
    # able to differentiate by display name using headerContains rules.
    domain_all_folders: dict[str, dict[str, int]] = {}
    for (domain, folder), count in domain_folder_counts.items():
        domain_all_folders.setdefault(domain, {})[folder] = count

    candidates = []
    header_candidates = []
    skipped_ambiguous = []
    for domain, folder_counts in domain_all_folders.items():
        total = sum(folder_counts.values())
        if total < min_occurrences:
            continue

        if len(folder_counts) > 1:
            # Domain maps to multiple folders — try display name splitting.
            # Group display names by folder to see if names are consistent.
            name_folders: dict[str, dict[str, int]] = {}
            for (d, name, folder), count in display_name_counts.items():
                if d == domain:
                    name_folders.setdefault(name, {})[folder] = count

            resolved_names = []
            still_ambiguous = []
            for name, nf_counts in name_folders.items():
                name_total = sum(nf_counts.values())
                if name_total < min_occurrences:
                    continue
                if len(nf_counts) == 1:
                    # This display name consistently maps to one folder
                    resolved_names.append({
                        "display_name": name,
                        "domain": domain,
                        "folder": next(iter(nf_counts)),
                        "count": name_total,
                    })
                else:
                    still_ambiguous.append({"name": name, "folders": nf_counts})

            if resolved_names:
                header_candidates.extend(resolved_names)
                logger.info(
                    f"  Domain '{domain}' is ambiguous by domain but {len(resolved_names)} "
                    f"display name(s) resolve cleanly → headerContains rules"
                )
                if still_ambiguous:
                    for sa in still_ambiguous:
                        folders_str = ", ".join(f"{f}: {c}" for f, c in sa["folders"].items())
                        logger.info(f"    Still ambiguous: \"{sa['name']}\" → {folders_str}")
            else:
                skipped_ambiguous.append({
                    "domain": domain,
                    "folders": folder_counts,
                    "sender_name": domain_examples.get(domain, domain),
                })
            continue

        # Single folder — safe to create a domain-level senderContains rule
        folder = next(iter(folder_counts))
        candidates.append({
            "domain": domain, "folder": folder, "count": total,
            "sender_name": domain_examples.get(domain, domain),
        })

    if skipped_ambiguous:
        logger.info(
            f"Skipped {len(skipped_ambiguous)} ambiguous domain(s) "
            f"(split across multiple folders — needs model judgment):"
        )
        for s in skipped_ambiguous:
            folders_str = ", ".join(f"{f}: {c}" for f, c in s["folders"].items())
            logger.info(f"  {s['domain']}: {folders_str}")

    if not candidates and not header_candidates:
        logger.info("No rule candidates found (need 2+ emails from same domain consistently to one folder)")
        return {"proposed": 0, "created": 0, "updated": 0, "candidates": [],
                "header_candidates": [], "ambiguous": skipped_ambiguous}

    # Read existing rules — we consolidate into one rule per folder
    existing_rules = _list_rules()
    folders_info = {f["name"].lower(): f for f in _list_folders()}

    # Build a map of existing rules. Three categories:
    #
    # 1. AUTO-SORT rules (named "Auto-sort: <Folder>") — domain-only,
    #    senderContains with no other conditions. The pipeline owns these
    #    and appends new domains to them.
    #
    # 2. HEADER rules — use headerContains to match sender display names.
    #    The pipeline creates these for ambiguous domains where different
    #    display names map to different folders (e.g., "Kroger Friday Deals"
    #    vs "Kroger").
    #
    # 3. SPECIFIC rules (everything else) — may combine fromAddresses +
    #    subjectContains or other conditions. The pipeline never touches
    #    these, but skips any domains/addresses they already cover.
    AUTO_RULE_PREFIX = "Auto-sort: "
    existing_auto_rules: dict[str, dict] = {}
    all_ruled_domains: set[str] = set()
    all_ruled_headers: set[str] = set()

    for rule in existing_rules:
        conditions = rule.get("conditions", {})

        # Collect domains from senderContains (auto-sort rules)
        domains_in_rule = [s.lower().strip() for s in conditions.get("senderContains", [])]
        all_ruled_domains.update(domains_in_rule)

        # Collect display names from headerContains rules
        headers_in_rule = [h.lower().strip() for h in conditions.get("headerContains", [])]
        all_ruled_headers.update(headers_in_rule)

        # Collect domains from specific rules that use fromAddresses with
        # additional conditions (subjectContains, bodyContains, etc.).
        # These are hand-crafted rules — the pipeline must never create a
        # domain-only senderContains rule that would override them.
        has_extra_conditions = bool(
            conditions.get("subjectContains")
            or conditions.get("bodyContains")
            or conditions.get("sentOnlyToMe")
            or conditions.get("notSentToMe")
            or conditions.get("headerContains")
        )
        if has_extra_conditions:
            for addr_obj in conditions.get("fromAddresses", []):
                addr = (addr_obj.get("emailAddress", {}).get("address", "")).lower()
                if "@" in addr:
                    domain = addr.split("@", 1)[1]
                    parts = domain.split(".")
                    parent = ".".join(parts[-2:]) if len(parts) > 2 else domain
                    all_ruled_domains.add(parent)
                    logger.debug(f"Domain {parent} covered by specific rule '{rule['name']}' — excluded from auto-sort")

        # Only track auto-sort rules for appending
        if rule["name"].startswith(AUTO_RULE_PREFIX):
            folder_name = rule["name"][len(AUTO_RULE_PREFIX):]
            existing_auto_rules[folder_name.lower()] = {
                "rule_id": rule["id"],
                "rule_name": rule["name"],
                "domains": domains_in_rule,
            }

    # Filter out domains that already appear in ANY rule
    new_candidates = [c for c in candidates if c["domain"] not in all_ruled_domains]
    # Filter out display names that already appear in headerContains rules
    new_header_candidates = [
        c for c in header_candidates
        if c["display_name"].lower() not in all_ruled_headers
    ]

    if not new_candidates and not new_header_candidates:
        total = len(candidates) + len(header_candidates)
        logger.info(f"All {total} candidate(s) already have rules")
        return {"proposed": total, "created": 0, "updated": 0,
                "already_exists": total, "candidates": candidates,
                "header_candidates": header_candidates,
                "ambiguous": skipped_ambiguous}

    # Group new domain candidates by target folder
    by_folder: dict[str, list[dict]] = {}
    for cand in new_candidates:
        by_folder.setdefault(cand["folder"], []).append(cand)

    # Group new header candidates by target folder
    header_by_folder: dict[str, list[dict]] = {}
    for cand in new_header_candidates:
        header_by_folder.setdefault(cand["folder"], []).append(cand)

    summary = {
        "proposed": len(new_candidates) + len(new_header_candidates),
        "candidates": new_candidates,
        "header_candidates": new_header_candidates,
        "created": 0,
        "updated": 0,
        "skipped": 0,
        "errors": [],
        "ambiguous": skipped_ambiguous,
    }

    # ── Create/update senderContains (domain) rules ──
    for folder_name, folder_cands in by_folder.items():
        new_domains = [c["domain"] for c in folder_cands]
        total_emails = sum(c["count"] for c in folder_cands)
        existing = existing_auto_rules.get(folder_name.lower())

        if existing:
            # Append domains to existing rule
            updated_domains = existing["domains"] + new_domains
            if dry_run:
                logger.info(
                    f"  Would update rule '{existing['rule_name']}': "
                    f"add {len(new_domains)} domain(s) ({total_emails} emails)"
                )
                continue
            try:
                _update_rule(existing["rule_id"], {
                    "conditions": {"senderContains": updated_domains},
                })
                logger.info(
                    f"  Updated rule '{existing['rule_name']}': "
                    f"added {', '.join(new_domains)}"
                )
                summary["updated"] += 1
            except Exception as e:
                logger.warning(f"  Failed to update rule for {folder_name}: {e}")
                summary["errors"].append(f"update {folder_name}: {e}")
                summary["skipped"] += 1
        else:
            # Create new rule for this folder with all its domains
            folder_info = folders_info.get(folder_name.lower())
            folder_id = folder_info["id"] if folder_info else _resolve_folder_id(folder_name)
            rule_name = f"{AUTO_RULE_PREFIX}{folder_name}"

            if dry_run:
                logger.info(
                    f"  Would create rule '{rule_name}': "
                    f"{len(new_domains)} domain(s) ({total_emails} emails)"
                )
                continue
            try:
                result = _create_rule(
                    rule_name,
                    {"senderContains": new_domains},
                    {"moveToFolder": folder_id, "stopProcessingRules": True},
                )
                logger.info(
                    f"  Created rule '{rule_name}' (ID: {result['id']}): "
                    f"{', '.join(new_domains)}"
                )
                summary["created"] += 1
            except Exception as e:
                logger.warning(f"  Failed to create rule for {folder_name}: {e}")
                summary["errors"].append(f"create {folder_name}: {e}")
                summary["skipped"] += 1

    # ── Create headerContains rules for ambiguous domains ──
    # These use the sender's display name (From: header) to differentiate
    # emails from the same domain that go to different folders.
    # One rule per folder, grouping all display names for that folder.
    for folder_name, folder_cands in header_by_folder.items():
        display_names = [c["display_name"] for c in folder_cands]
        total_emails = sum(c["count"] for c in folder_cands)
        domains = list(set(c["domain"] for c in folder_cands))

        folder_info = folders_info.get(folder_name.lower())
        folder_id = folder_info["id"] if folder_info else _resolve_folder_id(folder_name)
        rule_name = f"{domains[0]} → {folder_name} (by display name)"

        if dry_run:
            logger.info(
                f"  Would create headerContains rule '{rule_name}': "
                f"{display_names} ({total_emails} emails)"
            )
            continue
        try:
            result = _create_rule(
                rule_name,
                {"headerContains": display_names},
                {"moveToFolder": folder_id, "stopProcessingRules": True},
            )
            logger.info(
                f"  Created headerContains rule '{rule_name}' (ID: {result['id']}): "
                f"{', '.join(display_names)}"
            )
            summary["created"] += 1
        except Exception as e:
            logger.warning(f"  Failed to create headerContains rule for {folder_name}: {e}")
            summary["errors"].append(f"create headerContains {folder_name}: {e}")
            summary["skipped"] += 1

    # ── Remove ambiguous domains from senderContains rules ──
    # Domains that now have headerContains rules must be removed from
    # blanket senderContains rules, otherwise the domain rule fires first
    # (stopProcessingRules) and the display-name rules never match.
    header_domains = set(c["domain"] for c in new_header_candidates)
    for domain in header_domains:
        domain_upper = domain.upper()
        for folder_key, auto_rule in existing_auto_rules.items():
            if domain_upper in (d.upper() for d in auto_rule["domains"]):
                updated = [d for d in auto_rule["domains"]
                           if d.upper() != domain_upper]
                if dry_run:
                    logger.info(
                        f"  Would remove {domain} from '{auto_rule['rule_name']}'"
                    )
                    continue
                try:
                    if updated:
                        _update_rule(auto_rule["rule_id"], {
                            "conditions": {"senderContains": updated},
                        })
                        logger.info(
                            f"  Removed {domain} from '{auto_rule['rule_name']}'"
                        )
                    else:
                        _delete_rule(auto_rule["rule_id"])
                        logger.info(
                            f"  Deleted empty rule '{auto_rule['rule_name']}'"
                        )
                    summary["updated"] += 1
                except Exception as e:
                    logger.warning(
                        f"  Failed to remove {domain} from auto-sort: {e}"
                    )
                    summary["errors"].append(f"remove {domain}: {e}")

    if dry_run:
        summary["dry_run"] = True

    save_stage("05_rules.json", summary)
    return summary


# ── Main Pipeline ─────────────────────────────────────────────────────────

def _resolve_endpoint_and_key(args) -> tuple[str, str]:
    """Resolve endpoint URL and API key.

    Priority:
    1. CLI --api-key / --endpoint overrides
    2. GROQ_API_KEY env var
    3. Odysseus database — look up the endpoint by matching the URL or
       model prefix, decrypt the stored API key with the app's Fernet key.
    """
    endpoint = args.endpoint
    api_key = args.api_key or os.environ.get("GROQ_API_KEY", "")

    if api_key:
        return endpoint, api_key

    # Read from Odysseus database — avoids the fastapi import chain by
    # going straight to sqlite + cryptography (both always available).
    try:
        import sqlite3
        from cryptography.fernet import Fernet

        db_path = DATA_DIR / "app.db"
        key_path = DATA_DIR / ".app_key"
        if not db_path.exists() or not key_path.exists():
            return endpoint, ""

        conn = sqlite3.connect(str(db_path))
        # Find the endpoint whose base_url matches what we're calling
        rows = conn.execute(
            "SELECT base_url, api_key FROM model_endpoints WHERE is_enabled = 1"
        ).fetchall()
        conn.close()

        # Match by URL substring (e.g. "groq.com" in the endpoint)
        fernet = Fernet(key_path.read_bytes())
        for base_url, enc_key in rows:
            if not enc_key:
                continue
            # Check if this endpoint matches our target
            if base_url and any(part in endpoint for part in [base_url.split("//")[-1].split("/")[0]]):
                # Decrypt
                if enc_key.startswith("enc:"):
                    api_key = fernet.decrypt(enc_key[4:].encode("ascii")).decode("utf-8")
                else:
                    api_key = enc_key
                logger.info(f"Loaded API key from Odysseus DB (endpoint: {base_url})")
                return endpoint, api_key
    except Exception as e:
        logger.warning(f"Could not read API key from Odysseus DB: {e}")

    return endpoint, ""


def run_pipeline(args):
    """Run the full pipeline or individual stages."""
    model1 = args.model1
    model2 = args.model2
    endpoint, api_key = _resolve_endpoint_and_key(args)
    args.api_key = api_key

    # Auto-pause after N runs — pauses the scheduled task in the DB
    # so the cron stops firing entirely. Re-activate from the Tasks panel
    # or set status back to "active" to start a new blitz.
    if args.max_runs > 0:
        run_count_file = PIPELINE_DIR / ".run_count"
        PIPELINE_DIR.mkdir(parents=True, exist_ok=True)
        current_runs = int(run_count_file.read_text().strip()) if run_count_file.exists() else 0
        if current_runs >= args.max_runs:
            logger.info(f"Max runs reached ({current_runs}/{args.max_runs}) — pausing task.")
            run_count_file.write_text("0")
            # Pause the Blitz task in the database so cron stops
            try:
                import sqlite3
                db_path = DATA_DIR / "app.db"
                conn = sqlite3.connect(str(db_path))
                conn.execute(
                    "UPDATE scheduled_tasks SET status = 'paused' "
                    "WHERE name = 'Outlook Blitz' AND status = 'active'"
                )
                conn.commit()
                conn.close()
                logger.info("Outlook Blitz task paused. Re-activate from Tasks panel to run again.")
            except Exception as e:
                logger.warning(f"Could not pause task in DB: {e}")
            return
        run_count_file.write_text(str(current_runs + 1))
        logger.info(f"Blitz run {current_runs + 1}/{args.max_runs}")

    # Ensure OUTLOOK_CLIENT_ID is set
    if not os.environ.get("OUTLOOK_CLIENT_ID"):
        # Try loading from token cache
        cache_path = DATA_DIR / ".outlook_token_cache.json"
        if cache_path.exists():
            cache_data = json.loads(cache_path.read_text())
            for key, entry in cache_data.get("AccessToken", {}).items():
                cid = entry.get("client_id")
                if cid:
                    os.environ["OUTLOOK_CLIENT_ID"] = cid
                    break
        if not os.environ.get("OUTLOOK_CLIENT_ID"):
            logger.error("OUTLOOK_CLIENT_ID not set and no token cache found")
            sys.exit(1)

    # ── Stage 1: Fetch ──
    if args.resume:
        logger.info("Resuming from saved stage files...")
        emails = load_stage("01_emails.json")
    else:
        emails = fetch_emails(batch_size=args.batch_size)
        if not emails:
            logger.info("No emails in inbox. Nothing to do.")
            return
        save_stage("01_emails.json", emails)

    if args.fetch_only:
        logger.info(f"Fetch complete. {len(emails)} emails saved to {PIPELINE_DIR}/01_emails.json")
        return

    # ── Stage 2: Categorize ──
    if args.resume and (PIPELINE_DIR / "02_categories.json").exists():
        categories = load_stage("02_categories.json")
        logger.info(f"Loaded {len(categories)} cached categories")
    else:
        prompt = build_categorize_prompt(emails)
        response = llm_call(endpoint, model1, CATEGORIZE_SYSTEM, prompt,
                           temperature=0.3, max_tokens=4096, api_key=args.api_key)
        save_stage("02_categorize_raw.txt", response)
        categories = parse_categories(response, len(emails))
        save_stage("02_categories.json", categories)

    # ── Stage 3: Consolidate ──
    # Reuse existing folder map if available — consolidation only needs to
    # run once to establish the structure. Subsequent batches map new
    # categories into existing folders (1 RPD saved per batch).
    if (PIPELINE_DIR / "03_folder_map.json").exists():
        folder_map = load_stage("03_folder_map.json")
        logger.info(f"Reusing folder map: {folder_map.get('folders', [])}")
        # Map any new categories not in the existing mapping to the closest
        # folder by adding them with a default. The model only runs if there
        # are unmapped categories.
        existing_cats = set(folder_map.get("mapping", {}).keys())
        new_cats = [c["category"] for c in categories if c["category"] not in existing_cats]
        if new_cats and not args.skip_consolidate:
            logger.info(f"Found {len(new_cats)} new categories not in existing map — running consolidation...")
            # Pause between model calls to respect rate limits
            logger.info("Pausing 3s between model calls (rate limit headroom)...")
            time.sleep(3)
            # Ask model to map ONLY the new categories into existing folders
            # (or suggest new ones if nothing fits)
            _remap_prompt = (
                f"Existing folder structure: {json.dumps(folder_map['folders'])}\n\n"
                f"New categories that need to be mapped into these folders:\n"
                + "\n".join(f"  - {c}" for c in new_cats)
                + "\n\nMap each new category to the best existing folder. "
                "If a category truly doesn't fit any folder, you may add ONE new folder.\n\n"
                "Output JSON: {\"mapping\": {\"New Category\": \"Existing Folder\", ...}, "
                "\"new_folders\": []} (new_folders is empty if no new folder needed)"
            )
            response = llm_call(endpoint, model2, CONSOLIDATE_SYSTEM, _remap_prompt,
                               temperature=0.2, max_tokens=2048, api_key=args.api_key)
            import re as _re2
            _match = _re2.search(r"\{.*\}", response, _re2.DOTALL)
            if _match:
                remap = json.loads(_match.group())
                for cat, folder in remap.get("mapping", {}).items():
                    folder_map["mapping"][cat] = folder
                for nf in remap.get("new_folders", []):
                    if nf not in folder_map["folders"]:
                        folder_map["folders"].append(nf)
                save_stage("03_folder_map.json", folder_map)
                logger.info(f"Updated folder map with {len(new_cats)} new categories")
        elif new_cats:
            # --skip-consolidate: just map unknowns to Personal
            for cat in new_cats:
                folder_map["mapping"][cat] = "Personal"
            logger.info(f"Mapped {len(new_cats)} new categories to Personal (skip-consolidate)")
    else:
        # First run — full consolidation
        # Pause between model calls to respect rate limits
        logger.info("Pausing 3s between model calls (rate limit headroom)...")
        time.sleep(3)
        prompt = build_consolidate_prompt(categories)
        response = llm_call(endpoint, model2, CONSOLIDATE_SYSTEM, prompt,
                           temperature=0.2, max_tokens=2048, api_key=args.api_key)
        save_stage("03_consolidate_raw.txt", response)
        folder_map = parse_consolidation(response)
        save_stage("03_folder_map.json", folder_map)

    # Print the proposed structure for review
    print("\n" + "=" * 60)
    print("PROPOSED FOLDER STRUCTURE")
    print("=" * 60)
    if folder_map.get("reasoning"):
        print(f"\nRationale: {folder_map['reasoning']}")
    print(f"\nFolders ({len(folder_map['folders'])}):")
    for f in folder_map["folders"]:
        # Count emails mapping to this folder
        count = sum(1 for v in folder_map["mapping"].values() if v == f)
        cat_list = [k for k, v in folder_map["mapping"].items() if v == f]
        print(f"  > {f} ({count} categories)")
        for cat in cat_list[:8]:
            print(f"      - {cat}")
        if len(cat_list) > 8:
            print(f"      ... and {len(cat_list) - 8} more")
    print("=" * 60)

    # ── Stage 4: Move ──
    if args.dry_run:
        summary = apply_moves(emails, categories, folder_map, dry_run=True)
        save_stage("04_summary.json", summary)
        print(f"\nDRY RUN complete. Review the folder structure above.")
        print(f"To apply: python {__file__} --resume")
    else:
        # Confirm before moving
        if not args.yes:
            answer = input("\nApply this folder structure? [y/N] ").strip().lower()
            if answer not in ("y", "yes"):
                logger.info("Aborted by user. Run with --resume to retry.")
                return

        summary = apply_moves(emails, categories, folder_map, dry_run=False)
        save_stage("04_summary.json", summary)

        print(f"\nDone: Moved {summary.get('moved', 0)} emails into {len(summary.get('results', {}))} folders")
        if summary.get("failed"):
            print(f"Failed: {summary['failed']}")
        for folder, count in sorted(summary.get("results", {}).items()):
            print(f"  {folder}: {count}")

    # ── Stage 5: Learn Rules ──
    if not args.no_rules:
        print("\n" + "-" * 60)
        print("RULE LEARNING")
        print("-" * 60)
        rule_summary = learn_rules(
            emails, categories, folder_map,
            dry_run=args.dry_run,
            min_occurrences=args.min_rule_hits,
        )
        if rule_summary.get("proposed", 0) == 0:
            print("No new rule candidates (need 2+ emails from same domain consistently to one folder).")
        else:
            already = rule_summary.get("already_exists", 0)
            if already:
                print(f"  {already} domains already covered by existing rules")

            # Group by folder for cleaner output
            by_folder: dict[str, list] = {}
            for cand in rule_summary.get("candidates", []):
                by_folder.setdefault(cand["folder"], []).append(cand)
            for folder, cands in sorted(by_folder.items()):
                domains = ", ".join(c["domain"] for c in cands)
                total = sum(c["count"] for c in cands)
                print(f"  {folder}: +{len(cands)} domain(s) [{domains}] ({total} emails)")

            if not args.dry_run:
                if rule_summary["created"]:
                    print(f"\n  Created {rule_summary['created']} new rule(s)")
                if rule_summary["updated"]:
                    print(f"  Updated {rule_summary['updated']} existing rule(s)")
                if rule_summary.get("skipped"):
                    print(f"  Skipped {rule_summary['skipped']} (errors)")

        # Show ambiguous domains that were skipped
        ambiguous = rule_summary.get("ambiguous", [])
        if ambiguous:
            print(f"\n  Ambiguous domains (split across folders - model sorts these each run):")
            for a in ambiguous:
                folders_str = ", ".join(f"{f}: {c}" for f, c in a["folders"].items())
                print(f"    {a['domain']}: {folders_str}")
        print("-" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="Two-model email categorization pipeline for Outlook"
    )
    parser.add_argument("--batch-size", type=int, default=100,
                       help="Emails per batch (default: 100)")
    parser.add_argument("--endpoint", default=DEFAULT_ENDPOINT,
                       help=f"LLM endpoint URL (default: {DEFAULT_ENDPOINT})")
    parser.add_argument("--model1", default=DEFAULT_MODEL1,
                       help=f"Categorizer model (default: {DEFAULT_MODEL1})")
    parser.add_argument("--model2", default=DEFAULT_MODEL2,
                       help=f"Consolidator model (default: {DEFAULT_MODEL2})")
    parser.add_argument("--fetch-only", action="store_true",
                       help="Only fetch emails, don't categorize or move")
    parser.add_argument("--dry-run", action="store_true",
                       help="Run everything except actual moves")
    parser.add_argument("--resume", action="store_true",
                       help="Resume from saved JSON files instead of re-fetching")
    parser.add_argument("--api-key", default="",
                       help="API key for the LLM endpoint (or set GROQ_API_KEY env var)")
    parser.add_argument("--skip-consolidate", action="store_true",
                       help="Skip consolidation for new categories (map to Personal). Saves 1 API call.")
    parser.add_argument("--no-rules", action="store_true",
                       help="Skip rule learning stage")
    parser.add_argument("--min-rule-hits", type=int, default=2,
                       help="Minimum emails from a domain before creating a rule (default: 2)")
    parser.add_argument("--max-runs", type=int, default=0,
                       help="Auto-pause after N runs (tracks in data/outlook_pipeline/.run_count). 0 = unlimited.")
    parser.add_argument("--yes", "-y", action="store_true",
                       help="Skip confirmation prompt before moving")
    args = parser.parse_args()

    try:
        run_pipeline(args)
    except KeyboardInterrupt:
        print("\nInterrupted. Use --resume to continue from last stage.")
    except Exception as e:
        logger.exception(f"Pipeline failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
