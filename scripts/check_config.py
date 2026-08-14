#!/usr/bin/env python3
"""
check_config.py — report whether a usable config.json exists, without ever printing secrets.

Answers one question: can collect_roles.py run right now, or does someone need to supply
admin_key first?

"Usable" means the file exists, is valid JSON, and the admin_key field holds a real value
rather than the placeholder shipped in config.example.json. That last check matters because
the common failure mode isn't a missing file - it's a copied-but-never-edited one, which
looks fine to a `Path.exists()` check and then fails against the API with a confusing error
several steps later.

Prints a single JSON object to stdout and exits 0 whatever it finds (a missing config is a
normal state to report, not an error). Field *values* are never included in the output -
only booleans about them - so this is safe to run and quote in a chat transcript.

Usage:
    python3 check_config.py [--config /path/to/config.json]
"""

import argparse
import json
import sys
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
EXAMPLE_PATH = SKILL_DIR / "config.example.json"
REQUIRED_FIELDS = ("admin_key",)


def load_example() -> dict:
    """Placeholder values come from config.example.json itself, not a copy pasted in here.

    Keeping one source means editing the example can never silently stop a placeholder from
    being recognized as one.
    """
    try:
        return json.loads(EXAMPLE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {
            "admin_key": "sk-ant-admin01-REPLACE_ME",
        }


def candidate_paths(explicit):
    # Unannotated on purpose: `str | None` / `list[Path]` would make this file a syntax
    # error on Python 3.8, which is still what `python3` resolves to on some client machines.
    if explicit:
        return [Path(explicit)]
    # Working directory first: a config next to the user's own files is more likely to be
    # the one they just made than one sitting in an installed skill folder. Deduped because
    # the two collapse to one path whenever the skill folder *is* the working directory.
    paths = []
    for path in (Path.cwd() / "config.json", SKILL_DIR / "config.json"):
        if path not in paths:
            paths.append(path)
    return paths


def emit(payload: dict) -> None:
    print(json.dumps(payload, indent=2))
    sys.exit(0)


def main() -> None:
    parser = argparse.ArgumentParser(description="Check whether a usable config.json is present.")
    parser.add_argument("--config", default=None, help="Path to check (default: ./config.json, then <skill-dir>/config.json)")
    args = parser.parse_args()

    searched = candidate_paths(args.config)
    found = next((p for p in searched if p.is_file()), None)

    base = {
        "searched": [str(p) for p in searched],
        "example_path": str(EXAMPLE_PATH),
    }

    if found is None:
        emit({**base, "status": "missing_file", "config_path": None,
              "needs_from_user": list(REQUIRED_FIELDS)})

    try:
        config = json.loads(found.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        emit({**base, "status": "invalid_json", "config_path": str(found), "error": str(e),
              "needs_from_user": list(REQUIRED_FIELDS)})
    except OSError as e:
        emit({**base, "status": "unreadable", "config_path": str(found), "error": str(e),
              "needs_from_user": list(REQUIRED_FIELDS)})

    if not isinstance(config, dict):
        emit({**base, "status": "invalid_json", "config_path": str(found),
              "error": f"top-level JSON was {type(config).__name__}, expected an object",
              "needs_from_user": list(REQUIRED_FIELDS)})

    example = load_example()
    fields = {}
    needs = []
    for field in REQUIRED_FIELDS:
        value = config.get(field)
        present = isinstance(value, str) and value.strip() != ""
        is_placeholder = present and value.strip() == str(example.get(field, "")).strip()
        fields[field] = {"present": present, "is_placeholder": is_placeholder}
        if not present or is_placeholder:
            needs.append(field)

    emit({
        **base,
        "status": "ready" if not needs else "needs_values",
        "config_path": str(found),
        "fields": fields,
        "api_base_url_set": isinstance(config.get("api_base_url"), str) and config["api_base_url"].strip() != "",
        "needs_from_user": needs,
    })


if __name__ == "__main__":
    main()
