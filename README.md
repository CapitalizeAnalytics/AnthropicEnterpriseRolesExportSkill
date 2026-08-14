![Capitalize](capitalize-logo.png)

# Roles Export Collector (Claude Skill)

Pulls your organization's custom Claude roles into a set of detailed .xlsx workbooks you can review
in a spreadsheet, showing memberships, capabilities, permissions, connectors, and models for each role.

This is a **Claude Code skill** — a small bundle Claude reads and follows, so you don't run
any commands yourself. You don't need a config file before you start: if one isn't there,
Claude will ask you for the one value it needs, use it for the run, and then offer to
save it so you're not asked again.

## 1. Install it

Clone this repository into your Claude Code skills folder, picking the location based on how
broadly you want it available:

| Location | Available in |
|---|---|
| `<your-project>/.claude/skills/roles-export-collector/` | just that one Claude Code project |
| `%USERPROFILE%\.claude\skills\roles-export-collector\` (Windows) or `~/.claude/skills/roles-export-collector/` (Mac/Linux) | every Claude Code project you open |

```bash
git clone https://github.com/CapitalizeAnalytics/AnthropicEnterpriseRolesExportSkill.git ~/.claude/skills/roles-export-collector
```

On Windows PowerShell:

```powershell
git clone https://github.com/CapitalizeAnalytics/AnthropicEnterpriseRolesExportSkill.git "$env:USERPROFILE\.claude\skills\roles-export-collector"
```

Downloading the repo as a ZIP and extracting it to the same path works just as well. There's
no restart or registration step — Claude Code discovers skills by folder. Requires Python 3
and the `requests` and `openpyxl` packages:

```bash
pip install requests openpyxl
```

## 2. Ask Claude to run it

No slash command, no special syntax — just ask in plain language:

> Export our Claude roles to .xlsx files.

Claude finds this skill from what you asked for, checks whether a filled-in `config.json`
exists, and asks you for anything missing.

### The one value it will ask for

- `admin_key` — an **Admin API Key** from the Claude Console, with the format `sk-ant-admin01-...`.
  Create one at [claude.ai > Organization settings > API](https://claude.ai/admin-settings/api-access)
  → **Keys** → **Create key**. Select "Admin API Key" as the type.
  Full steps: https://platform.claude.com/docs/en/manage-claude/manage-users

### Prefer not to paste your key into chat?

Copy `config.example.json` to `config.json` in this folder and fill in the field yourself:

```json
{
  "admin_key": "sk-ant-admin01-REPLACE_ME",
  "api_base_url": "https://api.anthropic.com"
}
```

Leave `api_base_url` as-is unless you've been given a different endpoint. Then just ask
Claude for the export — it'll find the file and skip the questions.

`config.json` is listed in `.gitignore`, so a filled-in one won't be committed back to this
repository.

## What to expect

- Before running anything, Claude will show you a security warning and ask you to confirm
  — it won't run the script in the same response it shows that warning in. That's
  intentional: take it seriously the same way you would for any script that reads an API
  key, even one bundled with a tool you trust.
- Once you confirm, it makes API calls to fetch all roles and their permissions, then writes
  one .xlsx file per custom role to the `RoleDetails/` folder.
- Each workbook has tabs for:
  - **Membership** — groups assigned to the role and their members
  - **Capabilities** — every observed capability, with an Enabled/Disabled flag
  - **Permissions** — admin permissions with their access level
  - **Connectors** — connectors / tools / scopes with their approval setting
  - **Models** — models available to the role, with a Default flag
  - **Unclassified** — any permission rows the classifier didn't recognize (only written when non-empty)
- Afterward, if you supplied the key in chat, Claude offers to save it to `config.json`
  so future runs skip the questions. Saying no is fine — you'll just be asked again.

## Running it without Claude

The scripts are plain Python 3 and work on their own:

```bash
python3 scripts/check_config.py
python3 scripts/collect_roles.py --config config.json --output RoleDetails
```

Or without a config file at all, supplying the key through the environment:

```bash
ADMIN_API_KEY='<key>' python3 scripts/collect_roles.py --output RoleDetails
```

Add `--save-config config.json` to write that key to disk after a successful run. There
is intentionally no `--admin-key` flag: command-line arguments are visible in the process
table and shell history, so the key is only ever read from the environment or a config file.

## Troubleshooting

- **HTTP 403** — the key is missing scope or isn't an Admin API Key (they start with `sk-ant-admin01-`).
  Check that you created it at [claude.ai > Organization settings > API](https://claude.ai/admin-settings/api-access),
  not the Claude Console.
- **"No custom roles found"** — the organization doesn't have custom roles set up yet, or
  the API isn't returning them. Enterprise organizations with RBAC enabled should have at
  least the default roles visible. Confirm your organization is on Enterprise.
- **"Still using the config.example.json placeholder value"** — `config.json` was copied but
  never edited. Fill in the real `admin_key`.
- **"Requires: requests, openpyxl"** — install the dependencies with `pip install requests openpyxl`
  and try again.

## Security

Always review scripts before running them, especially ones that read API credentials. This
script is short and dependency-free specifically to make that review fast:

- **scripts/check_config.py** (~120 lines) — validates your config file without reading secrets
- **scripts/collect_roles.py** (~500 lines) — fetches roles via API and writes .xlsx files
- No obfuscation, no external dependencies beyond requests/openpyxl, no side effects beyond
  reading the config and writing output files

If you have an IT or security team, have them review these files before running the script.
If that's not available, you can ask Claude to review them with a prompt like:
"As a security specialist, analyze this script to identify any obfuscation, malicious intent,
or vulnerabilities which it would expose if run."
