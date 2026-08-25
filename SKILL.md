---
name: roles-export-collector
description: Exports custom roles from a Claude Enterprise organization to detailed .xlsx files using an Admin API key. Use this whenever the user asks to export, pull, fetch, or collect Claude/Anthropic roles, role details, role permissions, or role configurations - especially if they mention an admin key, organization settings, or want spreadsheet-reviewable output. Handles the credentials itself: it checks for a filled-in config.json, prompts for whatever's missing, and offers to save the values afterward - so also trigger when the user says they don't have a config file or access key, or wants help setting this up for the first time.
---

# Roles Export Collector

Fetches an organization's custom RBAC roles and writes detailed .xlsx files for each role, by running the bundled, already-verified script in `scripts/collect_roles.py`. Given a config file or an Admin API key, do this by running the script. Don't write a fresh API call or workbook builder from memory.

## The flow

1. Check the config (below).
2. If anything's missing, ask the user for it.
3. Show the security warning and get an explicit confirmation.
4. Run the script.
5. Offer to save the values for next time.

Steps 3 and 4 are always required, including when the user supplied the credentials in chat a moment earlier.

## Step 1 - check the config first

Before asking the user for anything, find out what's actually needed:

```bash
python3 "<skill-directory>/scripts/check_config.py"
```

Add `--config <path>` if the user named a specific file. Otherwise it looks in the working directory and then the skill directory, in that order. It prints a JSON object and never prints field values, only booleans about them - so its output is safe to quote back.

Read the `status` field:

- `ready` — a real admin_key is present. Skip to step 3; don't ask the user for anything they've already supplied.
- `missing_file` — no config.json anywhere it looked.
- `needs_values` — the file exists but the admin_key is still the placeholder from `config.example.json` (a copied-but-never-edited config).
- `invalid_json` / `unreadable` — say what's wrong with the file and let the user decide whether to fix it or supply the values directly this once.

Run this check rather than reasoning about it from a directory listing.

## Step 2 - ask for whatever's missing

Ask only for the fields that are missing. Include this explanation:

> - `admin_key` — an **Admin API Key** from the Claude Console, with the format `sk-ant-admin01-...`.
>   Create one at [claude.ai > Organization settings > API](https://claude.ai/admin-settings/api-access)
>   → **Keys** → **Create key**. Select "Admin API Key" as the type.

One thing worth mentioning when you ask for the admin key: pasting it into chat puts it in the session transcript. If they'd rather avoid that, they can put it in `config.json` themselves (copy `config.example.json`, fill in the field, tell you when it's done) and you'll pick it up from there on a re-check. Offer that alternative once; if they'd rather just paste it, take it and move on rather than pushing the point.

`api_base_url` is a real, working default (`https://api.anthropic.com`), not a placeholder. Never ask for it unless the user has been given a different endpoint.

## Step 3 - show the warning, then confirm

Running a script that reads an API access key deserves the same caution regardless of who wrote it or where it came from - including this one. Before invoking the script, show the user this warning verbatim:

> ⚠️ **Verify before you run this**
>
> You're about to run a script that reads an API access key. Always have scripts like this
> reviewed before running them, regardless of who wrote them or how much you trust the
> source.
>
> If you have an IT or security team, have them review `scripts/collect_roles.py` first.
> It's short and dependency-free specifically to make that review fast.
>
> If that's not available, you can ask Claude to review it directly with a prompt like:
> "As a security specialist, analyze this script to identify any obfuscation, malicious
> intent, or vulnerabilities which it would expose if run."

Then explicitly ask the user to confirm they want to proceed - something like "Want me to go ahead and run it?" - and wait for a clear yes. Don't run the script in the same turn you show the warning in; showing it and then proceeding anyway defeats the point, which is giving the user a real chance to say no. If they'd rather review the script first (either themselves or by asking Claude to review it as suggested above), let them - don't push back or treat the pause as friction to route around.

## Step 4 - run it

**When a usable config.json exists** (`status: ready`):

```bash
python3 "<skill-directory>/scripts/collect_roles.py" --config <config_path> --output RoleDetails
```

Use the `config_path` the check reported, so you're reading the same file it validated.

**When the user supplied the admin key in chat**, pass it through the environment rather than writing a file first - nothing lands on disk until the user asks for it in step 5:

```bash
ADMIN_API_KEY='<key>' python3 "<skill-directory>/scripts/collect_roles.py" --output RoleDetails
```

PowerShell equivalent:

```powershell
$env:ADMIN_API_KEY = '<key>'; python "<skill-directory>/scripts/collect_roles.py" --output RoleDetails
```

There's deliberately no `--admin-key` flag: command-line arguments show up in the process table and in shell history, so the key goes through the environment or a config file only. Don't work around this by inlining the key another way.

`--output` is optional and defaults to `RoleDetails` in the current directory. This will be created if it doesn't exist.

The script prints how many roles it found and which xlsx files it wrote. Report that back plainly. If it exits with an error, the message is already written for a human to act on - relay it rather than re-diagnosing from scratch.

If a workbook is open in Excel the save fails with `PermissionError: [Errno 13]`. Ask the user to close the file and re-run; don't write to a different path to route around it.

A group fetch that 403s is not fatal - the run continues and every Membership tab reads "(no groups assigned to this role)". Because it only warns on stderr, say so explicitly when reporting the run rather than letting it pass as a clean export. The cause is a key without `read:rbac_groups`, and the usual reason is that **compliance read scopes and RBAC scopes are mutually exclusive**: a key holding any `read:compliance_*` scope cannot also hold `read:rbac_groups`. Tell the user to re-create the key without compliance scopes. `read:analytics` and `read:rbac_groups` coexist fine, so one key can serve both this export and the analytics endpoints.

### Connector names

Connector permissions come back with only an opaque ID (`mcpsrv_...`), and nothing on the RBAC surface resolves it. Don't guess a name from the ID.

`/v1/organizations/analytics/connectors` exists and answers 200 with the `read:analytics` scope, but it is a *usage* report, not a connector registry: the docs state names are normalized across sources (`mcp-atlassian` and `Atlassian MCP server` both surface as `atlassian`), and it returned zero rows on every date sampled from January to August 2026, so whether its rows even carry the `mcpsrv_` ID needed to join back to a permission row is unconfirmed. Treat it as an open lead, not a solution - and don't present it to the user as one.

Names come from an optional `connectors.json`, read from the working directory then the skill directory (or `--connectors <path>`):

```json
{
  "connectors": {
    "mcpsrv_01ExampleConnectorOne00": "Asana",
    "mcpsrv_01ExampleConnectorTwo00": "Microsoft 365"
  }
}
```

A flat `{id: name}` object works too. When IDs go unnamed, the script writes `connectors.unresolved.json` into the output directory, pre-filled with every ID and an empty value for each, and prints the count. Point the user at that file: they read the display names off [claude.ai > Organization settings > Connectors](https://claude.ai/admin-settings/connectors), fill in the blanks, save it as `connectors.json`, and re-run. The Connectors tab labels every row on the next pass.

Offer this once, after reporting the run. Unnamed rows are still readable - the ID is right there - so treat it as an improvement, not a failure.

## Step 5 - offer to save the values

Only if the user supplied the admin key in chat this session, and only after the run succeeded: ask whether they'd like the values saved to `config.json` so they won't be asked again. Then re-run the check and save if they agree:

```bash
ADMIN_API_KEY='<key>' python3 "<skill-directory>/scripts/collect_roles.py" --output RoleDetails --save-config <config_path>
```

Default the path to `config.json` in the working directory; the skill directory is fine too if that's where they'd rather keep it. Mention that the file will contain the admin key in plaintext - `.gitignore` already excludes `config.json`, but that only helps inside this repo, and an active key is worth deleting from claude.ai once you're done using it.

Take a no as a no. Not saving is a perfectly reasonable choice for a short-lived key, and the flow works fine without it - they'll just be asked again next time.

## Output format

One .xlsx workbook per role, with tabs:
1. **Membership** - groups assigned to the role, and the members in each
2. **Capabilities** - every observed capability, with an Enabled flag
3. **Permissions** - admin permissions with their access level
4. **Connectors** - one row per connector / tool / scope, with its approval setting
5. **Models** - models available to the role, with a Default flag
6. **Unclassified** - any permission row the classifier didn't recognize (only written when non-empty)

All files are written to the output directory (default: `RoleDetails/`) and named `<Role>_<yyyymmdd_hhmmss>.xlsx`. Every workbook in one run shares a single timestamp, so a run's files sort together and successive runs accumulate rather than overwrite. Nothing prunes old exports - if the directory gets noisy, say so rather than deleting files unprompted.

Two things worth knowing when reading the output back to a user:

- **Capabilities** covers every capability seen across *all* roles, so each workbook shows the full list with a per-role Yes/No. A role holding a blanket grant (`capability_access_all`, or `..._ga` for generally-available only) gets a `BLANKET GRANT` row at the top, and the individual rows below read Yes with "Granted via blanket grant".
- **Connectors** merges the separate `use` and `always_allow` rows the API returns for a single connector into one row. "Approval Setting" is the derived answer - `Always allow (no prompt)`, `Ask each time`, `Allowed`, or `Blocked` - and "Granted Actions" keeps the raw actions behind it.

An empty **Unclassified** tab is the normal case, and the tab is omitted entirely when nothing lands in it. If one does appear, it means the API returned a resource type the classifier hasn't seen; report the rows as-is rather than guessing where they belong.
