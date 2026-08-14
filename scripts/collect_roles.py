#!/usr/bin/env python3
"""
collect_roles.py — Export Claude Enterprise organization roles to .xlsx files.

Resolves an admin_key (from a config JSON file or environment variable), calls the
Claude Roles API to fetch all custom RBAC roles, and writes detailed .xlsx files for
each role with tabs for:
    1. Membership   - groups assigned to the role, and the members in each
    2. Capabilities - every observed capability, with an Enabled flag
    3. Permissions  - admin permissions with their access level
    4. Connectors   - connectors / tools / scopes with their approval setting
    5. Models       - models available to the role, with a Default flag
    6. Unclassified - any permission row the classifier didn't recognize

Requires: requests, openpyxl (install with: pip install requests openpyxl)
Python 3 standard library for the rest.

Usage:
    python3 collect_roles.py --config /path/to/config.json [--output RoleDetails]
    ADMIN_API_KEY='...' python3 collect_roles.py [--output RoleDetails]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterator

try:
    import requests
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit(
        "This script requires 'requests' and 'openpyxl'.\n"
        "Install them with: pip install requests openpyxl"
    )

SKILL_DIR = Path(__file__).resolve().parent.parent
EXAMPLE_PATH = SKILL_DIR / "config.example.json"

BASE = "https://api.anthropic.com/v1/organizations"
BETA_HEADER = "ce-user-management-2026-07-13"

ENV_ADMIN_KEY = "ADMIN_API_KEY"
ENV_API_BASE_URL = "ROLES_API_BASE_URL"

DEFAULT_API_BASE_URL = "https://api.anthropic.com"

# Rate limit is 100 req/min org-wide. Stay under it.
MIN_INTERVAL = 0.65
_last_call = 0.0


# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

def load_example_placeholders() -> dict:
    """The placeholder values live in config.example.json, not duplicated here."""
    try:
        example = json.loads(EXAMPLE_PATH.read_text(encoding="utf-8"))
        if isinstance(example, dict):
            return example
    except (OSError, json.JSONDecodeError):
        pass
    return {
        "admin_key": "sk-ant-admin01-REPLACE_ME",
    }


def load_config_file(path: Path) -> dict:
    if not path.exists():
        sys.exit(
            f"Config file not found at '{path}'.\n"
            f"Either create it (copy config.example.json and fill in admin_key),\n"
            f"or supply the value via the {ENV_ADMIN_KEY} environment variable."
        )
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        sys.exit(f"'{path}' is not valid JSON: {e}")
    if not isinstance(config, dict):
        sys.exit(f"'{path}' must contain a JSON object at the top level.")
    return config


def resolve_settings_source(args) -> dict:
    """Merge config file and environment into one resolved config.

    Precedence, highest first: environment variables, config file. The override order
    lets someone run this once with values typed in a chat session without having to
    write a credential to disk first.

    Note: there is no --admin-key flag. Command-line arguments are visible to other
    users via the process table and land in shell history; the key is accepted only
    through the environment or the config file for that reason.
    """
    config = load_config_file(Path(args.config)) if args.config else {}

    admin_key = os.environ.get(ENV_ADMIN_KEY) or config.get("admin_key")
    api_base_url = (
        os.environ.get(ENV_API_BASE_URL)
        or config.get("api_base_url")
        or DEFAULT_API_BASE_URL
    )

    resolved = {
        "admin_key": (admin_key or "").strip(),
        "api_base_url": (api_base_url or "").strip(),
    }

    example = load_example_placeholders()
    for key, value in resolved.items():
        if not value or value == str(example.get(key, "")):
            sys.exit(
                f"Missing or placeholder value for '{key}'. "
                f"Set it in config.json, via the {ENV_ADMIN_KEY} environment variable, "
                f"or both."
            )

    return resolved


def save_config(config: dict, path: Path) -> None:
    """Write resolved config to disk (for reuse on next invocation)."""
    payload = {
        "admin_key": config.get("admin_key"),
        "api_base_url": config.get("api_base_url"),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


# --------------------------------------------------------------------------
# HTTP
# --------------------------------------------------------------------------

class Client:
    def __init__(self, api_key: str, api_base: str = DEFAULT_API_BASE_URL):
        self.api_base = api_base
        self.s = requests.Session()
        self.s.headers.update({
            "x-api-key": api_key,
            "anthropic-beta": BETA_HEADER,
        })

    def get(self, path: str, **params) -> dict:
        """GET with throttling and 429/5xx backoff."""
        global _last_call
        url = f"{self.api_base}/v1/organizations{path}"
        for attempt in range(6):
            gap = time.monotonic() - _last_call
            if gap < MIN_INTERVAL:
                time.sleep(MIN_INTERVAL - gap)
            _last_call = time.monotonic()

            r = self.s.get(url, params={k: v for k, v in params.items() if v is not None})
            if r.status_code == 429 or r.status_code >= 500:
                wait = float(r.headers.get("retry-after", 2 ** attempt))
                sys.stderr.write(f"  {r.status_code} on {path}, retrying in {wait:.0f}s\n")
                time.sleep(wait)
                continue
            if not r.ok:
                raise RuntimeError(f"{r.status_code} {path}: {r.text[:400]}")
            return r.json()
        raise RuntimeError(f"gave up after retries: {path}")

    def paginate(self, path: str, limit: int = 100) -> Iterator[dict]:
        """Cursor pagination: feed next_page back in as `page` until it's null."""
        page = None
        while True:
            body = self.get(path, limit=limit, page=page)
            yield from body.get("data", [])
            page = body.get("next_page")
            if not page:
                return


# --------------------------------------------------------------------------
# Classification
# --------------------------------------------------------------------------

CONNECTOR_RESOURCES = {"connector", "connector_tool", "connector_scope", "all_connectors"}
BLANKET_ACTIONS = {"capability_access_all", "capability_access_all_ga"}
ACCESS_RANK = {"no_access": 0, "view": 1, "can_view": 1, "manage": 2, "can_manage": 2}


def bucket(perm: dict) -> str:
    res = perm.get("resource") or {}
    rtype = res.get("type", "")
    action = perm.get("action", "")

    if rtype in CONNECTOR_RESOURCES:
        return "connectors"
    if "model" in rtype or action.startswith("model"):
        return "models"
    if rtype == "organization":
        if action.startswith("permission_"):
            return "permissions"
        if action in BLANKET_ACTIONS or action.startswith("capability_"):
            return "capabilities"
    return "unclassified"


def prettify(token: str) -> str:
    """capability_web_search -> Web Search"""
    for prefix in ("capability_", "permission_", "model_"):
        if token.startswith(prefix):
            token = token[len(prefix):]
            break
    return re.sub(r"[_\-]+", " ", token).strip().title()


# --------------------------------------------------------------------------
# Fetch
# --------------------------------------------------------------------------

def fetch_roles(c: Client) -> list[dict]:
    return list(c.paginate("/rbac_roles"))


def fetch_groups_by_role(c: Client) -> dict[str, list[dict]]:
    """Reverse-index groups onto the roles attached to them, members included."""
    index: dict[str, list[dict]] = defaultdict(list)
    for g in c.paginate("/rbac_groups"):
        roles = g.get("roles")
        if roles is None:
            # Degraded read, not an empty group. Retry once before trusting it.
            retry = c.get(f"/rbac_groups/{g['id']}")
            roles = retry.get("roles")
            if roles is None:
                sys.stderr.write(
                    f"  WARNING: group '{g.get('name')}' returned null roles twice; "
                    f"its role assignments are missing from this export\n"
                )
                continue
        if not roles:
            continue
        members = list(c.paginate(f"/rbac_groups/{g['id']}/members"))
        for role_id in roles:
            index[role_id].append({**g, "_members": members})
    return index


def fetch_permissions(c: Client, role_id: str) -> list[dict]:
    return list(c.paginate(f"/rbac_roles/{role_id}/permissions"))


# --------------------------------------------------------------------------
# Workbook
# --------------------------------------------------------------------------

FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=11)
BODY = Font(name=FONT, size=10)
GROUP_FONT = Font(name=FONT, bold=True, size=10)


def write_sheet(ws, headers: list[str], rows: list[list[Any]], bold_rows: set[int] = frozenset()):
    ws.append(headers)
    for cell in ws[1]:
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"

    for i, row in enumerate(rows):
        ws.append(row)
        font = GROUP_FONT if i in bold_rows else BODY
        for cell in ws[ws.max_row]:
            cell.font = font

    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    for col in range(1, len(headers) + 1):
        letter = get_column_letter(col)
        widest = max(
            [len(str(headers[col - 1]))] +
            [len(str(r[col - 1])) for r in rows if col - 1 < len(r) and r[col - 1] is not None]
        )
        ws.column_dimensions[letter].width = min(max(widest + 3, 12), 60)


def safe_name(name: str) -> str:
    cleaned = re.sub(r'[\\/*?:\[\]]', "-", name).strip() or "role"
    return cleaned[:100]


def build_workbook(role: dict, groups: list[dict], perms: list[dict],
                   universe: dict[str, set[str]], out_dir: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    buckets: dict[str, list[dict]] = defaultdict(list)
    for p in perms:
        buckets[bucket(p)].append(p)

    # --- 1. Membership -----------------------------------------------------
    rows, bold = [], set()
    for g in sorted(groups, key=lambda x: x.get("name", "")):
        members = g.get("_members", [])
        source = "Identity provider (SCIM)" if g.get("source_type") == "scim" else "Created in Claude"
        rows.append([g.get("name"), "", "", source, len(members)])
        bold.add(len(rows) - 1)
        for m in sorted(members, key=lambda x: (x.get("email") or "")):
            rows.append(["", m.get("email"), m.get("user_id"), "", ""])
    if not rows:
        rows = [["(no groups assigned to this role)", "", "", "", ""]]
    write_sheet(wb.create_sheet("Membership"),
                ["Group", "Member Email", "User ID", "Group Source", "Member Count"],
                rows, bold)

    # --- 2. Capabilities ---------------------------------------------------
    granted = {p["action"] for p in buckets["capabilities"]}
    blanket = granted & BLANKET_ACTIONS
    rows = []
    if blanket:
        label = "All capabilities" if "capability_access_all" in blanket else "All generally available capabilities"
        rows.append([label, "BLANKET GRANT", "Yes",
                     "Covers every capability in scope. Excludes model access and admin permissions."])
    for action in sorted(universe["capabilities"] - BLANKET_ACTIONS):
        enabled = "Yes" if (action in granted or blanket) else "No"
        note = "Granted via blanket grant" if (blanket and action not in granted) else ""
        rows.append([prettify(action), action, enabled, note])
    write_sheet(wb.create_sheet("Capabilities"),
                ["Capability", "Action Key", "Enabled", "Notes"], rows)

    # --- 3. Permissions ----------------------------------------------------
    by_action: dict[str, str] = {}
    for p in buckets["permissions"]:
        a, lvl = p["action"], str(p.get("access") or p.get("level") or "granted")
        if ACCESS_RANK.get(lvl, 1) >= ACCESS_RANK.get(by_action.get(a, ""), -1):
            by_action[a] = lvl
    rows = [[prettify(a), a, prettify(by_action.get(a, "no_access"))]
            for a in sorted(universe["permissions"])]
    write_sheet(wb.create_sheet("Permissions"),
                ["Admin Permission", "Action Key", "Access Type"], rows)

    # --- 4. Connectors -----------------------------------------------------
    rows = []
    for p in sorted(buckets["connectors"], key=lambda x: json.dumps(x.get("resource"), sort_keys=True)):
        res = p.get("resource") or {}
        rtype = res["type"]
        scope = {
            "all_connectors": "All connectors",
            "connector": "Whole connector",
            "connector_tool": "Individual tool",
            "connector_scope": "OAuth scope",
        }.get(rtype, rtype)
        target = res.get("tool_name") or res.get("scope") or res.get("connector_id") or "(all)"
        rows.append([res.get("connector_id", "-"), scope, target, prettify(p.get("action", ""))])
    if not rows:
        rows = [["(no connector grants - all connectors blocked for this role)", "", "", ""]]
    write_sheet(wb.create_sheet("Connectors"),
                ["Connector ID", "Scope", "Tool / Scope", "Approval Setting"], rows)

    # --- 5. Models ---------------------------------------------------------
    rows = []
    for p in buckets["models"]:
        res = p.get("resource") or {}
        model = res.get("model") or res.get("model_id") or res.get("name") or p.get("action", "")
        rows.append([
            model,
            "Yes",
            "Yes" if (res.get("is_default") or p.get("default")) else "",
            res.get("max_effort") or p.get("max_effort") or "",
        ])
    if not rows:
        rows = [["(no model rows returned - see README on reading model access)", "", "", ""]]
    write_sheet(wb.create_sheet("Models"),
                ["Model", "Available", "Default", "Max Effort Level"], rows)

    # --- 6. Unclassified ---------------------------------------------------
    if buckets["unclassified"]:
        rows = [[(p.get("resource") or {}).get("type", ""), p.get("action", ""),
                 json.dumps(p.get("resource"), sort_keys=True)]
                for p in buckets["unclassified"]]
        write_sheet(wb.create_sheet("Unclassified"),
                    ["Resource Type", "Action", "Raw Resource"], rows)

    path = out_dir / f"{safe_name(role.get('name', role['id']))}.xlsx"
    wb.save(path)
    return path


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser(
        description="Export Claude Enterprise organization roles to .xlsx files."
    )
    ap.add_argument(
        "--config",
        default=None,
        help="Path to config.json (default: ./config.json, then <skill-dir>/config.json)"
    )
    ap.add_argument(
        "--output",
        default="RoleDetails",
        help="Output directory for .xlsx files (default: RoleDetails)"
    )
    ap.add_argument(
        "--save-config",
        default=None,
        help="Save resolved config to this path for reuse (only via environment/CLI, not from existing config)"
    )
    ap.add_argument(
        "--no-groups",
        action="store_true",
        help="Skip the Membership tab (use if your key isn't scoped for all linked orgs)"
    )
    args = ap.parse_args()

    # Resolve configuration
    config = resolve_settings_source(args)
    api_base = config.get("api_base_url") or DEFAULT_API_BASE_URL

    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    c = Client(config["admin_key"], api_base)

    try:
        print("Fetching roles...")
        roles = fetch_roles(c)
        print(f"  {len(roles)} role(s)")
        if not roles:
            print("No custom roles found. Confirm the org is on Enterprise and has custom roles.")
            return 0

        groups_by_role: dict[str, list[dict]] = {}
        if not args.no_groups:
            print("Fetching groups and members...")
            try:
                groups_by_role = fetch_groups_by_role(c)
            except RuntimeError as e:
                sys.stderr.write(f"  Group fetch failed ({e}).\n"
                                 f"  Continuing without Membership data. "
                                 f"Groups need read:rbac_groups on an all-linked-orgs key.\n")

        print("Fetching permissions per role...")
        perms_by_role = {}
        for r in roles:
            perms_by_role[r["id"]] = fetch_permissions(c, r["id"])
            print(f"  {r.get('name')}: {len(perms_by_role[r['id']])} permission row(s)")

        # Observed universe across every role
        universe = {"capabilities": set(), "permissions": set()}
        for perms in perms_by_role.values():
            for p in perms:
                b = bucket(p)
                if b in universe:
                    universe[b].add(p["action"])

        print("Writing workbooks...")
        for r in roles:
            p = build_workbook(r, groups_by_role.get(r["id"], []),
                               perms_by_role[r["id"]], universe, out)
            print(f"  {p}")

        # Save config if requested
        if args.save_config:
            config_path = Path(args.save_config)
            save_config(config, config_path)
            print(f"\nConfig saved to {config_path}")

        print(f"\nDone. {len(roles)} role(s) exported to {out}")
        return 0

    except RuntimeError as e:
        sys.stderr.write(f"Error: {e}\n")
        return 1


if __name__ == "__main__":
    sys.exit(main())
