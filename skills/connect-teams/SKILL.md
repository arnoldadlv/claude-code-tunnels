---
name: connect-teams
description: "Skill for connecting a Microsoft Teams channel to the Claude-Code-Tunnels Orchestrator. Guides Azure Bot registration, configures Bot Framework webhook, collects credentials, and verifies the connection. Execute with /claude-code-tunnels:connect-teams."
---

# Connect Teams Channel

Adds a Microsoft Teams channel to an existing Claude-Code-Tunnels Orchestrator.
Connects via Bot Framework webhook (aiohttp HTTP server), so a public URL or tunnel (e.g. ngrok) is required.

## Rules

- **Never proceed without asking the user**
- **Auto-detected values are presented as numbered choices first** — the user only needs to enter a number
- If an existing credentials file is found, always confirm with the user before overwriting
- Credential files use the `key : value` format (spaces on both sides of the colon)
- The ARCHIVE/ directory must not be committed to git

---

## Step 0: Environment Preflight (CRITICAL)

Connecting Teams requires the orchestrator to be installed, pip packages present, and a credentials file.
Check each item in order — **if any check fails, do not proceed to the next step until it is resolved.**

### 0-1. Verify orchestrator.yaml

**Why it is needed**: the Teams adapter reads channel activation status, port, and the ARCHIVE path from orchestrator.yaml. Without this file, adapter initialization will fail.

```bash
if [ ! -f "orchestrator.yaml" ]; then
  echo "orchestrator.yaml not found."
  echo "Please run /claude-code-tunnels:setup-orchestrator first."
  # -> stop here
fi
```

### 0-2. Verify ARCHIVE_PATH

**Why it is needed**: Teams credentials (app_id, app_password) are stored in `ARCHIVE/teams/credentials`.

```bash
ARCHIVE_PATH=$(python3 -c "import yaml; print(yaml.safe_load(open('orchestrator.yaml')).get('archive', 'ARCHIVE'))")
```

```
Confirming credential storage path.

  [1] $ARCHIVE_PATH   <- value read from orchestrator.yaml
  [2] Enter manually

Number:
```

### 0-3. Verify pip + packages

**Why they are needed**: `botbuilder-integration-aiohttp` provides the Bot Framework adapter. Without it, `from botbuilder.core import ...` will raise an ImportError.

```bash
$PYTHON_CMD -c "import botbuilder.core" 2>/dev/null
```

If not installed:
```
The following package required for Teams connection is not installed:
  - botbuilder-integration-aiohttp  (Bot Framework adapter and aiohttp integration)

  [1] Install now ($PIP_CMD install botbuilder-integration-aiohttp)
  [2] Skip (install manually and continue)

Number:
```

### 0-4. Check for existing credentials

**Why it is needed**: if Teams is already configured, the user must decide whether to overwrite the existing setup.

```bash
if [ -f "$ARCHIVE_PATH/teams/credentials" ]; then
  echo "Existing credentials found"
fi
```

If an existing file is found:
```
Existing Teams credentials already exist:
  app_id:       xxxxxxxx-xxxx-...
  app_password: ****

  [1] Overwrite (enter new values)
  [2] Keep existing values (update configuration only)
  [3] Cancel

Number:
```

---

## Step 1: Azure Bot Setup Guide

If no Azure Bot exists yet, provide guidance:

```
Azure Bot Registration:
1. Go to https://portal.azure.com → Create a resource → "Azure Bot"
2. Fill in:
   - Bot handle: choose a unique name
   - Subscription / Resource group: select or create
   - Type of App: Single Tenant or Multi Tenant
   - Creation type: "Create new Microsoft App ID"
3. After creation, go to the Bot resource → Configuration
   - Copy the Microsoft App ID (this is your app_id)
   - Click "Manage Password" → New client secret → copy the value (this is your app_password)
4. Under Channels → Add Microsoft Teams channel
5. Set the Messaging endpoint to your public URL:
   - Example: https://your-domain.com/api/messages
   - For local dev: use ngrok — ngrok http 3978 — then use the https URL

Are you ready? (yes — start entering credentials / no — show detailed guide)
```

---

## Step 2: Collect Credentials (2 fields)

**Ask the user for each field one at a time. Empty values are not accepted.**

```
─────────────────────────────────────────────────────────────────
1. app_id
   The Microsoft App ID from your Azure Bot resource → Configuration.
   Format: UUID (e.g. 12345678-abcd-1234-efgh-123456789012)
   Enter:

2. app_password
   The client secret you created under "Manage Password".
   This is shown only once when created — if lost, generate a new one.
   Enter:
─────────────────────────────────────────────────────────────────
```

Summary after collection:
```
Teams Credentials entered:
  app_id:       12345678-abcd-1234-efgh-123456789012
  app_password: ****

Save with these values? (yes/no)
```

---

## Step 3: Save Configuration

After user confirmation:

```bash
mkdir -p $ARCHIVE_PATH/teams/

cat > $ARCHIVE_PATH/teams/credentials << 'EOF'
app_id : $APP_ID
app_password : $APP_PASSWORD
EOF
```

Update orchestrator.yaml:
```yaml
channels:
  teams:
    enabled: true
    port: 3978
```

---

## Step 4: Connection Test

```bash
cd $PROJECT_ROOT && ./start-orchestrator.sh --fg &
sleep 5
# Confirm "Teams channel started on port 3978" in the logs
```

- Success → "Teams connection complete. @mention the bot in a Teams channel to test it."
- Failure → show the error log to the user and analyze the cause. Do not retry automatically.

**Note**: The bot must be reachable at a public URL for Teams to deliver messages.
For local development, use a tunnel like ngrok: `ngrok http 3978`

## Credential File Format

```
app_id : 12345678-abcd-1234-efgh-123456789012
app_password : your-client-secret-value
```
