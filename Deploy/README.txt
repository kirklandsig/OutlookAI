OutlookAI - Deployment Guide
============================

Three install shapes are supported:
  A. Single workstation (developer or power user)
  B. Multi-user RDS / Terminal Server (primary deployment target)
  C. IT-managed image / silent install

All three share the same installer:
  Deploy/Install-OutlookAI.ps1

The installer requires Administrator. It is idempotent: running it again
upgrades in place.


PREREQUISITES (all shapes)
---------------------------
- Windows 10 / 11 (Pro or Enterprise) or Windows Server 2019 / 2022 / 2025.
- Microsoft Outlook desktop (2016, 2019, 2021, 2024, or Microsoft 365).
- .NET Framework 4.7.2 or later.
- Visual Studio Tools for Office Runtime:
  https://aka.ms/VSTORuntime
- Microsoft Edge WebView2 Evergreen Runtime. The installer will install it
  silently if missing. To pre-stage offline, vendor the bootstrapper:
    .\Deploy\Fetch-WebView2Bootstrapper.ps1


WHAT THE INSTALLER DOES
-----------------------
1. Backs up any v1 config to:
     C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>
2. Cleans up stale Outlook add-in registrations for every user profile on
   the machine (ClickOnce subscription, VSTA, VSTO SolutionMetadata,
   Inclusion list, Add/Remove Programs, Outlook AddInLoadTimes, ClickOnce
   app cache). This prevents AddInAlreadyInstalledException on upgrade.
3. Closes any running Outlook.exe.
4. Copies the published build to:
     C:\Program Files\OutlookAI
5. Writes a fresh v2 config.xml at C:\Program Files\OutlookAI\config.xml
   if none exists. Server-authoritative defaults: Model, CodexAuthPath.
6. Creates the shared OAuth credential directory at:
     C:\ProgramData\OutlookAI
   with Authenticated Users: Modify (RDS shared-credential model).
7. Renames any per-user %APPDATA%\OutlookAI\config.xml that lacks the v2
   CodexAuthPath element to <name>.v1.backup.<timestamp>.
8. Configures VSTO trust + Inclusion list (HKLM, 64-bit + WOW6432Node).
9. Registers OutlookAI for all users.
10. Configures the Default User profile so new RDS users auto-load it.


SHAPE A - SINGLE WORKSTATION
-----------------------------
Use case: developer machine, power user, single-user laptop or desktop.

Steps:

1. Clone the repo or download the latest Release zip.

2. Publish a Release build:

   & "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
     "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release `
     /p:Platform="Any CPU" /p:PublishDir="C:\OutlookAI\"

3. Run the installer elevated:

   Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
   .\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

4. Open Outlook -> click AI Assistant on the ribbon.

5. Open a compose window or use the taskpane button -> click any action ->
   the default browser opens the ChatGPT OAuth consent page (the consent
   screen says "Codex CLI" because OutlookAI reuses the public Codex
   client_id; this is expected).

6. Sign in. The browser confirms and OutlookAI returns to its taskpane.

7. Verification:
     SHA256 of C:\Program Files\OutlookAI\OutlookAI.dll matches the
     SHA256 of the staged C:\OutlookAI\OutlookAI.dll.


SHAPE B - MULTI-USER RDS / TERMINAL SERVER
-------------------------------------------
Use case: shared server where many interactive users open Outlook with
their own profile and you want the same ChatGPT credential to back all of
them.

Steps:

1. From an admin workstation, build a Release publish bundle and copy it
   to the RDS server, e.g. to C:\OutlookAI on the server.

2. On the RDS server, in an elevated PowerShell:

   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
   cd C:\OutlookAI
   .\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

3. Designate one admin user as the "first run" user. Have them log in,
   open Outlook, open AI Assistant, and click any action. The OAuth
   browser flow runs once; the resulting auth.json is shared with every
   user on this server.


ACCEPTED RISK - SHARED OAuth CREDENTIAL
----------------------------------------
auth.json sits in C:\ProgramData\OutlookAI with Authenticated Users:
Modify. Any signed-in interactive user on this server can:
  - Read auth.json and copy tokens off the box.
  - Use those tokens to call OpenAI directly until revoked.
  - Delete or corrupt auth.json (signs everyone out).
  - Replace auth.json with their own ChatGPT tokens (other users'
    traffic then bills to the attacker's ChatGPT account and the
    attacker can observe every call).

Only deploy this build to RDS servers where every interactive user is
trusted with the ChatGPT account that signs in.


ROTATING CREDENTIALS
--------------------
Phase 1 has no remote revocation; rotation is a manual two-step:

1. On the RDS server, as any user who knows the OutlookAI admin password:
   - Open Outlook -> AI Assistant -> gear icon (Settings).
   - Enter the OutlookAI admin password.
   - Click "Sign Out" in the ChatGPT Account section.
   - Click "Sign In" and authenticate with the new ChatGPT account.

2. From the OpenAI side (recommended after any suspected leak):
   - Sign the previous account out at https://chatgpt.com/#settings
     (rotates session tokens server-side).


SHAPE C - IT-MANAGED IMAGE / SILENT INSTALL
--------------------------------------------
Use case: corporate Windows image, MDT/SCCM rollout, or any deployment
where the install must complete without operator interaction.

Pre-stage the WebView2 runtime so the installer has no internet dependency
during image build:

   .\Deploy\Fetch-WebView2Bootstrapper.ps1

Then bake the installer into the image:

1. Copy the published OutlookAI bundle (the contents of the PublishDir)
   to a known location on the gold image, e.g. C:\OutlookAI.

2. Add a run-once install task (Group Policy startup script, SCCM task
   sequence, MDT package, or Task Scheduler "At startup") that calls:

   powershell.exe -NoProfile -ExecutionPolicy Bypass `
     -File C:\OutlookAI\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

   The script enforces #Requires -RunAsAdministrator. Schedule the task
   to run as SYSTEM or a local admin account.

3. Post-install verification (run on a sample VM provisioned from the
   image):

   $staged    = (Get-FileHash 'C:\OutlookAI\OutlookAI.dll' -Algorithm SHA256).Hash
   $installed = (Get-FileHash 'C:\Program Files\OutlookAI\OutlookAI.dll' -Algorithm SHA256).Hash
   if ($staged -ne $installed) { throw "OutlookAI DLL hash mismatch" }

4. First-run experience for end users: open Outlook -> AI Assistant ->
   first action triggers the OAuth flow per user (or, on RDS-style images,
   relies on the shared auth.json from Shape B).


VERIFICATION (all shapes)
--------------------------
1. Outlook -> File -> Options -> Add-ins lists "OutlookAI".
2. Outlook ribbon shows the AI Assistant group.
3. Click AI Assistant -> taskpane opens.
4. Click a Quick Action / type a chat message -> response streams in.
5. C:\ProgramData\OutlookAI\auth.json exists (after first sign-in).
6. (Optional) SHA256 of installed OutlookAI.dll matches staged build.


TROUBLESHOOTING
---------------

Add-in shows in list but won't load / keeps unchecking:
  1. Confirm VSTO Runtime is installed (Programs and Features:
     "Microsoft Visual Studio 2010 Tools for Office Runtime").
  2. File -> Options -> Add-ins -> Manage: Disabled Items -> Go ->
     enable OutlookAI if listed there.
  3. File -> Options -> Add-ins -> Manage: COM Add-ins -> Go ->
     tick OutlookAI; note any error.
  4. Event Viewer -> Windows Logs -> Application -> look for "Outlook"
     or ".NET Runtime" errors.

OAuth sign-in doesn't open a browser:
  - Confirm the user has a default browser configured.
  - Confirm http://localhost:1455 is not blocked locally; the installer
    does not modify firewall rules because the listener is loopback only.

Sign-in returns immediately with "OAuth state mismatch":
  - Click Sign In again; this is usually a stale browser tab racing the
    fresh authorize URL.

ChatGPT Codex backend returns 4xx:
  - 401: token rotated remotely; click Sign Out then Sign In.
  - 429: ChatGPT subscription rate-limited; wait or upgrade plan.
  - 403 with HTML body: Cloudflare challenge; retry once.

Realtime voice fails with "beta_api_shape_disabled":
  - Should not happen on this build (no OpenAI-Beta header is sent).
    If you see it, file an issue with the exact error_id.

PDF export fails with HRESULT 0x8007139F:
  - Indicates a WebView2 user-data folder conflict. The current build
    isolates the PDF renderer at
    C:\Users\<user>\AppData\Local\OutlookAI\WebView2PdfData. If conflicts
    persist, delete that folder and retry.

Search takes very long on large mailboxes:
  - The current build caps interactive broad all-mail scans at 200
    folders and early-stops once enough candidates have been collected.
    See VSTO2\OutlookAI\Services\Tools\SearchFallbackBudget.cs.


UNINSTALL
---------
1. Open PowerShell as Administrator.
2. Run: .\Uninstall-OutlookAI.ps1

This removes:
  - HKLM Outlook add-in registration (64-bit + WOW6432Node).
  - C:\Program Files\OutlookAI install directory.
  - C:\ProgramData\OutlookAI\auth.json + sidecar refresh lock.

Backups under C:\ProgramData\OutlookAI\Backups are intentionally preserved
for rollback.


ROLLBACK TO v1
--------------
1. Run Uninstall-OutlookAI.ps1.
2. Reinstall the v1.x publish artifacts using the v1 installer.
3. Restore the latest backup over the new install:

   Copy-Item `
     "C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>" `
     "C:\Program Files\OutlookAI\config.xml" -Force


SUPPORT
-------
For issues, file a GitHub issue at:
  https://github.com/kirklandsig/OutlookAI/issues
