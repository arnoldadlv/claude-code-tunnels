---
name: setup-orchestrator
description: "Skill for installing Claude-Code-Tunnels (Project Orchestrator). Installs the PO into the current project directory, registers workspaces, and connects messenger channels — all in one run. Execute with /claude-code-tunnels:setup-orchestrator."
---

# Claude-Code-Tunnels Setup

Installs Claude-Code-Tunnels (Project Orchestrator) into the user's project directory.
The PO discovers sub-projects/workspaces and delegates tasks via `claude-agent-sdk query(cwd=workspace/)`.

## Plugin Source

This plugin directory contains the orchestrator code and templates.
`PLUGIN_DIR` = the directory two levels above this SKILL.md (the plugin root).

Included files: `orchestrator/`, `templates/`, `orchestrator.yaml`, `install.sh`, `requirements.txt`

## Rules

- **Never proceed without asking the user** — all environment variables and paths must be confirmed by the user
- However, **auto-detected values are presented as numbered choices first** — the user only needs to enter a number
- Do not modify the original code logic. Only change paths and configuration
- Preserve existing CLAUDE.md content. Append only when necessary
- The ARCHIVE/ directory must never be committed to git

---

## Phase 0: Environment Preflight (CRITICAL — must run first)

The Orchestrator depends on the Python runtime, pip, and the Claude SDK.
Installation paths and versions vary by system, so the current environment must be verified against requirements before installation.
**If any check fails → do not proceed to the next step until it is resolved.**

### 0-1. Python Runtime

**Why it is needed**: the entire orchestrator is written in Python, and `claude-agent-sdk` requires Python 3.10+.

Auto-detect available Python installations on the system:

```bash
candidates=()
for cmd in python3 python python3.12 python3.11 python3.10; do
  full=$(command -v "$cmd" 2>/dev/null) || continue
  ver=$("$full" -c "import sys; v=sys.version_info; print(f'{v.major}.{v.minor}')" 2>/dev/null) || continue
  major=${ver%%.*}; minor=${ver#*.}
  if [ "$major" -ge 3 ] && [ "$minor" -ge 10 ]; then
    candidates+=("$full ($cmd $ver)")
  fi
done
```

Present the detected results as **numbered choices**:

```
Select a Python 3.10+ runtime.
The Orchestrator code and claude-agent-sdk will run using this Python.

  [1] /usr/bin/python3       (python3 3.11.5)   <- detected
  [2] /usr/local/bin/python3.12 (python3.12 3.12.1) <- detected
  [3] Enter manually

Number or path:
```

- User enters `1` or `2` → set that path as `PYTHON_CMD`
- User enters `3` or a direct path → verify that path's version and use it
- Zero candidates detected → "Could not find Python 3.10+. Please enter the path manually."

### 0-2. pip

**Why it is needed**: used to install and verify Python packages such as claude-agent-sdk, aiohttp, and pyyaml.

```bash
pip_candidates=()
for cmd in "$PYTHON_CMD -m pip" pip3 pip; do
  if $cmd --version &>/dev/null 2>&1; then
    pip_candidates+=("$cmd")
  fi
done
```

```
Select pip. It will be used to install dependency packages.

  [1] /usr/bin/python3 -m pip  (pip 23.2.1)   <- detected
  [2] Enter manually

Number or path:
```

### 0-3. Claude Code CLI

**Why it is needed**: `claude-agent-sdk` internally calls the `claude` CLI binary. Without it, `query()` calls will fail.

```bash
claude_path=$(command -v claude 2>/dev/null)
```

- Found → "Claude CLI detected: `$claude_path` (`claude --version`). OK?"
- Not found →
  ```
  Claude CLI not found.
  The claude-agent-sdk query() internally calls the claude binary, so it is required.

    [1] Enter path manually (e.g. ~/.npm/bin/claude)
    [2] Install now (npm install -g @anthropic-ai/claude-code)
    [3] Install later (continue without installing — errors may occur at runtime)

  Number:
  ```

### 0-4. Required Python Packages

**Why they are needed**: each package serves a distinct role.
- `claude-agent-sdk`: the core library for invoking Claude Code programmatically from Python
- `aiohttp`: used by channel adapters (Telegram polling) and the remote listener to handle async HTTP
- `pyyaml`: parses the orchestrator.yaml configuration file

```bash
declare -A pkg_status
for pkg in claude_agent_sdk aiohttp yaml; do
  if $PYTHON_CMD -c "import $pkg" 2>/dev/null; then
    pkg_status[$pkg]="OK"
  else
    pkg_status[$pkg]="NOT INSTALLED"
  fi
done
```

If any packages are missing:
```
The following packages are not installed:
  - claude-agent-sdk (Claude SDK — orchestrator cannot run without it)
  - aiohttp          (async HTTP — channel/remote connections will fail without it)

Install command: $PIP_CMD install claude-agent-sdk aiohttp pyyaml

  [1] Install now
  [2] Skip (install manually later)

Number:
```

### Preflight Results

```
Environment Preflight Results
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Python:            /usr/bin/python3 (3.11.5)         ✓
  pip:               /usr/bin/python3 -m pip (23.2.1)  ✓
  Claude CLI:        /usr/local/bin/claude (1.0.35)    ✓
  claude-agent-sdk:  0.3.0                             ✓
  aiohttp:           3.9.1                             ✓
  pyyaml:            6.0.1                             ✓
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
All checks passed. Proceed with setup? (yes/no)
```

---

## Phase 1: Collect User Input

Now that the environment check is complete, determine the installation location and channels.
Auto-detect possible values and **present them as numbered choices**. The user only needs to enter a number.

### 1-1. PROJECT_ROOT

Present the current directory and its parent as candidates:

```bash
cwd=$(pwd)
parent=$(dirname "$cwd")

# Treat paths with 2 or more subdirectories as project root candidates
candidates=()
for d in "$cwd" "$parent"; do
  subdir_count=$(find "$d" -maxdepth 1 -mindepth 1 -type d ! -name '.*' | wc -l)
  if [ "$subdir_count" -ge 1 ]; then
    candidates+=("$d  (${subdir_count} subdirectories)")
  fi
done
```

```
Select the project root directory.
Sub-projects/workspaces must reside inside this directory.

  [1] /home/user/my-projects  (4 subdirectories)  <- current location
  [2] /home/user              (7 subdirectories)  <- parent
  [3] Enter manually

Number or absolute path:
```

Validation: `test -d` && `test -w`

### 1-2. ARCHIVE_PATH

```
Select the credential storage path.
Slack/Telegram tokens and similar secrets will be stored here. (excluded from git)

  [1] $PROJECT_ROOT/ARCHIVE   <- recommended default
  [2] Enter manually

Number or absolute path:
```

### 1-3. CHANNELS

```
Select the messenger channels to connect.
The Orchestrator will receive messages from these channels and execute tasks.
Multiple selections allowed — separate numbers with commas (e.g. 1,3)

  [1] slack      — Slack Socket Mode (no public IP required)
  [2] telegram   — Telegram Bot long polling (no public IP required)
  [3] teams      — Microsoft Teams Bot Framework webhook (requires public URL or tunnel)
  [4] Configure later (skip)

Number:
```

### Input Validation (mandatory)

```bash
test -d "$PROJECT_ROOT" || echo "ERROR: $PROJECT_ROOT does not exist."
test -w "$PROJECT_ROOT" || echo "ERROR: $PROJECT_ROOT is not writable."
mkdir -p "$ARCHIVE_PATH" 2>/dev/null || echo "ERROR: Failed to create $ARCHIVE_PATH."
```

**If validation fails → re-present choices for that item only. Do not automatically substitute a value.**

---

## Phase 2: Copy Orchestrator Code

Show the list of files to be copied and confirm with the user:

```bash
PLUGIN_DIR="<plugin root path>"

echo "The following files will be copied to $PROJECT_ROOT:"
echo "  orchestrator/    <- Python package (PO, executor, router, channels)"
echo "  .claude/rules/   <- delegation, task-log, notification rules"
echo "  start-orchestrator.sh <- startup script (uses $PYTHON_CMD)"
echo "Proceed? (yes/no)"

# After user confirmation:
cp -r $PLUGIN_DIR/orchestrator/ $PROJECT_ROOT/orchestrator/
mkdir -p $PROJECT_ROOT/.claude/rules/
cp $PLUGIN_DIR/templates/rules/*.md $PROJECT_ROOT/.claude/rules/

# start-orchestrator.sh — apply PYTHON_CMD
sed "s|python3|$PYTHON_CMD|g" $PLUGIN_DIR/templates/start-orchestrator.sh.template \
  > $PROJECT_ROOT/start-orchestrator.sh
chmod +x $PROJECT_ROOT/start-orchestrator.sh
```

---

## Phase 3: Generate orchestrator.yaml

Generate using the collected user inputs, then display the content and confirm:

```yaml
root: $PROJECT_ROOT
archive: $ARCHIVE_PATH
channels:
  slack:
    enabled: true/false
  telegram:
    enabled: true/false
  teams:
    enabled: true/false
    port: 3978
remote_workspaces: []
```

---

## Phase 4: CLAUDE.md Configuration

- Does not exist → generate from `$PLUGIN_DIR/templates/CLAUDE.md.template`
- Exists but has no Orchestrator mention → append the orchestrator section
- Already contains Orchestrator content → skip

---

## Phase 5: Workspace Discovery

Auto-detect subdirectories and present as a checklist:

```bash
ls $PROJECT_ROOT/   # exclude: orchestrator/, ARCHIVE/, .tasks/, .claude/, .git/, hidden folders
```

```
The following subdirectories were found.
Select the ones to register as workspaces (separate numbers with commas, or enter all):

  [1] project-a/        (CLAUDE.md present)
  [2] project-b/        (CLAUDE.md present)
  [3] project-c/        (no CLAUDE.md — a default will be generated)
  [4] data-scripts/     (no CLAUDE.md — a default will be generated)

Number (e.g. 1,2,3 or all):
```

Workspaces without a CLAUDE.md will be informed that a default one will be created.

---

## Phase 6: Channel Configuration

Run the corresponding channel skill in sequence based on the selected channels:
- Slack → `/claude-code-tunnels:connect-slack`
- Telegram → `/claude-code-tunnels:connect-telegram`
- Teams → `/claude-code-tunnels:connect-teams`

Dependency installation (show list to user and confirm):
```
Packages to install:
  Base:   claude-agent-sdk aiohttp pyyaml
  Slack:  slack-bolt slack-sdk
  Teams:  botbuilder-integration-aiohttp

Command: $PIP_CMD install claude-agent-sdk aiohttp pyyaml slack-bolt slack-sdk botbuilder-integration-aiohttp

  [1] Install now
  [2] Skip (install manually)

Number:
```

---

## Phase 7: Test & Complete

```bash
cd $PROJECT_ROOT && ./start-orchestrator.sh --fg &
sleep 3
# Slack: confirm "Socket Mode" in logs
# Telegram: confirm bot username in logs
```

**If the test fails → explain the situation to the user with the error message. Do not retry automatically.**

Final summary:
```
Setup Complete!
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Project Root:  /home/user/my-projects
  Archive:       /home/user/my-projects/ARCHIVE
  Python:        /usr/bin/python3 (3.11.5)
  Channels:      slack ✓, telegram ✗
  Workspaces:    project-a, project-b, project-c

Files created:
  orchestrator/          <- Python package
  orchestrator.yaml      <- configuration file
  start-orchestrator.sh  <- startup script
  .claude/rules/         <- delegation rules
  CLAUDE.md              <- PO description

Next steps:
  ./start-orchestrator.sh              <- start in background
  ./start-orchestrator.sh --fg         <- foreground (debug)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

## Credential File Format

All credential files share the same format:
```
key : value
```
Spaces are required on both sides of the colon. Even if bot_token contains a colon, only the first ` : ` is used as the split point.

## File Structure (after installation)

```
PROJECT_ROOT/
├── orchestrator/          <- PO, executor, router, channels
│   ├── __init__.py
│   ├── main.py
│   ├── server.py
│   ├── po.py
│   ├── executor.py
│   ├── router.py
│   ├── channel/           <- Slack, Telegram adapters
│   └── remote/            <- listener, deploy helpers
├── orchestrator.yaml      <- configuration file
├── start-orchestrator.sh  <- startup script
├── CLAUDE.md              <- project description for the PO
├── .claude/rules/         <- delegation, task-log, notification rules
├── ARCHIVE/               <- credentials (excluded from git)
├── .tasks/                <- task history logs
├── project-a/             <- workspace
│   └── CLAUDE.md
└── project-b/             <- workspace
    └── CLAUDE.md
```
