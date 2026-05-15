OutlookAI - RDS Deployment Guide (v2 - ChatGPT OAuth)
======================================================

PREREQUISITES
-------------
- Windows Server 2022 or 2025
- .NET Framework 4.7.2 (or later)
- Visual Studio Tools for Office Runtime
  Download: https://aka.ms/VSTORuntime


WHAT v2 CHANGES
---------------
- Anthropic / Claude API key removed
- Separate OpenAI Whisper API key removed
- Single ChatGPT OAuth sign-in shared per server
- Text generation -> https://chatgpt.com/backend-api/codex/responses
- Voice transcription -> wss://api.openai.com/v1/realtime
- Both billed against the user's ChatGPT consumer subscription


INSTALLATION
------------
1. Publish from Visual Studio to C:\OutlookAI
   (or copy published files there)

2. Copy Install-OutlookAI.ps1 to C:\OutlookAI

3. Open PowerShell as Administrator

4. Enable script execution (one-time):
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

5. If script is blocked, unblock it:
   Unblock-File -Path "C:\OutlookAI\Install-OutlookAI.ps1"

6. Run the install script:
   cd C:\OutlookAI
   .\Install-OutlookAI.ps1

The script always installs to C:\Program Files\OutlookAI. The -InstallPath
parameter from v1 has been removed because manifest URLs and registry
entries are pinned to that path.


WHAT THE INSTALLER DOES
-----------------------
1. Backs up any existing v1 config to:
     C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>
2. Replaces C:\Program Files\OutlookAI with the new build
3. Writes a fresh v2 config.xml (server-authoritative defaults)
4. Creates the shared OAuth folder at C:\ProgramData\OutlookAI
5. Grants Authenticated Users: Modify on that folder (see ACCEPTED RISK)
6. Renames any per-user %APPDATA%\OutlookAI\config.xml that lacks the
   v2 <CodexAuthPath> element to <name>.v1.backup.<timestamp>
7. Configures VSTO trust + Inclusion List (HKLM, 64-bit + WOW6432Node)
8. Registers the add-in for all users
9. Configures the Default User profile so new users auto-load it


FIRST USE (per server)
----------------------
1. Have a designated admin user log into the RDS server
2. Open Outlook -> open a compose window -> click AI Assistant
3. Click any text action (e.g., Proofread) or the mic button
4. The default browser opens an OpenAI sign-in page (consent shows
   "Codex CLI" -- this is expected; OutlookAI reuses the public Codex
   OAuth client_id in Phase 1)
5. Sign in with the ChatGPT account you want all users to share
6. The browser confirms sign-in and OutlookAI returns to its task pane
7. The session is now persisted in C:\ProgramData\OutlookAI\auth.json
   and is shared by every user on this RDS server


ACCEPTED RISK: SHARED OAuth CREDENTIAL
--------------------------------------
auth.json sits in C:\ProgramData\OutlookAI with Authenticated Users: Modify.

This means any signed-in user on this RDS server can:
- Read auth.json and copy the OAuth tokens off the box
- Use those tokens to call OpenAI directly until they are revoked
- Delete or corrupt auth.json (every user is signed back out)
- Replace auth.json with their own ChatGPT account's tokens (every other
  user's traffic then bills to the attacker's ChatGPT account, and the
  attacker can see everything OutlookAI sends through the API)

Only deploy this build to RDS servers where every interactive user is
trusted with the ChatGPT account that signs in.

If the user base or trust posture changes, follow ROTATING CREDENTIALS
below and reconsider whether to switch ACL to a dedicated local group.


ROTATING CREDENTIALS
--------------------
Phase 1 has no remote revocation; rotation is a manual two-step:

1. On the RDS server (any user with admin pwd):
   - Open Outlook -> AI Assistant -> gear icon (Settings)
   - Enter the OutlookAI admin password
   - Click "Sign Out" in the ChatGPT Account section
   - Click "Sign In" and authenticate with the new account
   This deletes the local auth.json and writes a fresh one with the
   new ChatGPT account's tokens.

2. From the OpenAI side (recommended after any suspected leak):
   - Sign the previous account out at https://chatgpt.com/#settings
     (rotates session tokens server-side)


VERIFICATION
------------
1. Open Outlook -> File -> Options -> Add-ins -> OutlookAI is listed
2. Open a compose window -> AI Assistant button on the ribbon
3. Quick Action (Proofread / Revise / etc.) returns text via the
   ChatGPT Codex backend
4. Mic button records and returns a transcript via the Realtime
   WebSocket


UNINSTALLATION
--------------
1. Open PowerShell as Administrator
2. Run: .\Uninstall-OutlookAI.ps1

This removes:
- HKLM Outlook add-in registration (64-bit + WOW6432Node)
- C:\Program Files\OutlookAI install directory
- C:\ProgramData\OutlookAI\auth.json + sidecar refresh lock

Backups under C:\ProgramData\OutlookAI\Backups are intentionally
preserved for rollback.


ROLLBACK TO v1
--------------
1. Run Uninstall-OutlookAI.ps1
2. Reinstall the v1.x publish artifacts using the v1 installer
3. Restore the latest backup over the new install:
   Copy-Item `
     "C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>" `
     "C:\Program Files\OutlookAI\config.xml" -Force


TROUBLESHOOTING
---------------

Add-in shows in list but won't load / keeps unchecking:
1. Confirm VSTO Runtime is installed (Programs and Features:
   "Microsoft Visual Studio 2010 Tools for Office Runtime")
2. File -> Options -> Add-ins -> Manage: Disabled Items -> Go ->
   if OutlookAI is listed, Enable it
3. File -> Options -> Add-ins -> Manage: COM Add-ins -> Go ->
   tick OutlookAI; note any error
4. Event Viewer -> Windows Logs -> Application -> look for "Outlook"
   or ".NET Runtime" errors

OAuth sign-in doesn't open a browser:
- Confirm the user has a default browser configured
- Confirm http://localhost:1455 is not blocked locally; the installer
  does not modify firewall rules because the listener is loopback only

Sign-in returns immediately with "OAuth state mismatch":
- Click Sign In again; this is usually a stale browser tab racing the
  fresh authorize URL

ChatGPT Codex backend returns 4xx:
- Status code 401: token rotated remotely; click Sign Out then Sign In
- Status code 429: ChatGPT subscription rate-limited; wait or upgrade plan
- Status code 403 with HTML body: Cloudflare challenge; retry once

Realtime voice fails with "beta_api_shape_disabled":
- Should not happen on this build (we never send OpenAI-Beta header).
  If you see it, file an issue with the exact error_id.


SUPPORT
-------
For issues, contact your IT administrator.
