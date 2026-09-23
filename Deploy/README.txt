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
- Microsoft Edge WebView2 Evergreen Runtime. If it is missing, the
  installer runs the WebView2 bootstrapper from the bundle, which downloads
  the runtime from Microsoft. On a machine without internet access,
  install the runtime first with Microsoft's Evergreen Standalone
  Installer.


WHAT THE INSTALLER DOES
-----------------------
1. Backs up the existing config.xml to:
     C:\ProgramData\OutlookAI\Backups\config.xml.v1.backup.<timestamp>
2. Cleans up stale Outlook add-in registrations for every user profile on
   the machine (ClickOnce subscription, VSTA, VSTO SolutionMetadata,
   Inclusion list, Add/Remove Programs, Outlook AddInLoadTimes, ClickOnce
   app cache). This prevents AddInAlreadyInstalledException on upgrade.
3. Closes any running Outlook.exe.
4. Copies the published build to:
     C:\Program Files\OutlookAI
5. Keeps C:\Program Files\OutlookAI\config.xml as it is on an update, so
   hand-set values (Model, VoiceModel, MaxBulkExportRows,
   ModelCatalogClientVersion, ...) survive; a missing CodexAuthPath is
   added. The file is never removed along with the old build, so an install
   that stops partway leaves it in place for the next run. A file that
   isn't valid XML is left for you to fix (the installer warns). A fresh
   install, or a v1 (Claude-era) file, gets the v2 template: the
   AdminPassword carried over, plus CodexAuthPath. It sets no Model, so the
   default is the top model in the ChatGPT model list.
6. Creates the shared folder for the sign-in, Settings and model list:
     C:\ProgramData\OutlookAI
   with Authenticated Users: Modify (RDS shared-credential model), and
   warns if that permission can't be set.
7. Renames any per-user v1 (Claude-era) %APPDATA%\OutlookAI\config.xml to
   <name>.v1.backup.<timestamp>. Per-user files that Settings saved before
   v2.2.2 are kept, but the server-wide C:\ProgramData\OutlookAI\config.xml
   wins over them (see SETTINGS below).
8. Configures VSTO trust + Inclusion list (HKLM, 64-bit + WOW6432Node).
9. Registers OutlookAI for all users.
10. Configures the Default User profile so new RDS users auto-load it.


GET THE INSTALL BUNDLE
----------------------
Every release on https://github.com/kirklandsig/OutlookAI/releases has an
install bundle, OutlookAI-vX.Y.Z-RDS-Deploy.zip, and its .sha256. The zip
holds the published add-in, Install-OutlookAI.ps1, Uninstall-OutlookAI.ps1,
the WebView2 bootstrapper, this guide and version.json. Download both
files, then in PowerShell, in the download folder:

   $zip = ".\OutlookAI-vX.Y.Z-RDS-Deploy.zip"
   (Get-FileHash $zip -Algorithm SHA256).Hash -eq (Get-Content "$zip.sha256").Trim()
   Unblock-File $zip
   Expand-Archive $zip -DestinationPath C:\OutlookAI

The hash check must print True. The steps below assume C:\OutlookAI.

To build the bundle from source instead, see "Contributing" in the
repository's README.md (Deploy\Make-ReleaseZip.ps1 writes the same zip).


SHAPE A - SINGLE WORKSTATION
-----------------------------
Use case: developer machine, power user, single-user laptop or desktop.

Steps:

1. Get the install bundle (above) into C:\OutlookAI.

2. Run the installer elevated:

   Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
   C:\OutlookAI\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

3. Open Outlook -> click AI Assistant on the ribbon.

4. Open a compose window or use the taskpane button -> click any action ->
   the default browser opens the ChatGPT OAuth consent page (the consent
   screen says "Codex CLI" because OutlookAI reuses the public Codex
   client_id; this is expected).

5. Sign in. The browser confirms and OutlookAI returns to its taskpane.

6. Verification:
     SHA256 of C:\Program Files\OutlookAI\OutlookAI.dll matches the
     SHA256 of C:\OutlookAI\OutlookAI.dll from the bundle.


SHAPE B - MULTI-USER RDS / TERMINAL SERVER
-------------------------------------------
Use case: shared server where many interactive users open Outlook with
their own profile and you want the same ChatGPT credential to back all of
them.

Steps:

1. Get the install bundle (above) onto the RDS server, e.g. into
   C:\OutlookAI on the server.

2. On the RDS server, in an elevated PowerShell:

   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine
   cd C:\OutlookAI
   .\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

3. Designate one admin user as the "first run" user. Have them log in,
   open Outlook, open AI Assistant, and click any action. The OAuth
   browser flow runs once; the resulting auth.json is shared with every
   user on this server.

Exports: where Folder Redirection points Documents at a network share,
Excel/PDF exports go to %LOCALAPPDATA%\OutlookAI\Reports\ instead of
Documents\OutlookAI\Reports\.

Later versions install from Settings -> Updates (see UPDATES below).


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
  - Edit models.json or config.xml there. config.xml holds the Settings
    every user gets, including the admin password (in plain text), the
    model and which write tools are on.

Only deploy this build to RDS servers where every interactive user is
trusted with the ChatGPT account that signs in.


SETTINGS (ALL USERS)
--------------------
Settings (gear icon, admin password) saves the model, reasoning effort,
write tools and admin password for every user on the machine, in:

  C:\ProgramData\OutlookAI\config.xml

Each user's Outlook reads it the next time it starts. Settings opens on
what is saved there now and saves only what you change, so two admins
don't overwrite each other's changes.

After the first Settings save that file holds all five of these settings,
so to change one by hand, edit it there: the same setting in
C:\Program Files\OutlookAI\config.xml only applies while the ProgramData
file doesn't have it. The Program Files file is the only place for
CodexAuthPath, VoiceModel, MaxBulkExportRows and ModelCatalogClientVersion
(see SERVER SETTINGS below).

Before v2.2.2, Settings also saved a copy per user in
%APPDATA%\OutlookAI\config.xml, which kept overriding later changes for
that user. Now the ProgramData file wins over those copies; a copy only
fills in a setting the ProgramData file doesn't have. If Settings rejects
your admin password after updating, use the one in
C:\ProgramData\OutlookAI\config.xml.

Going back to a version before v2.2.2 makes the per-user copies override
again: delete them first (%APPDATA%\OutlookAI\config.xml for each user).


SERVER SETTINGS (Program Files config.xml)
------------------------------------------
C:\Program Files\OutlookAI\config.xml is only edited by hand; updates keep
it. Besides fallbacks for the Settings above, it holds settings the
Settings dialog doesn't show:

  CodexAuthPath              where the shared ChatGPT sign-in is stored
                             (default C:\ProgramData\OutlookAI\auth.json)
  VoiceModel                 the voice transcription model
                             (default gpt-realtime-1.5)
  MaxBulkExportRows          the most messages one complete-list Excel
                             export collects (default 2000, at most 10000)
  ModelCatalogClientVersion  the Codex version the model list request
                             identifies as (blank = latest; see MODEL LIST)

Each user's Outlook reads it the next time it starts.


MODEL LIST (models.json)
------------------------
The model dropdown and each model's reasoning efforts come from the model
catalog ChatGPT publishes for the signed-in account. Settings -> AI
Behavior -> Update Models fetches it and caches it to:

  C:\ProgramData\OutlookAI\models.json        (shared, every user)
  %LOCALAPPDATA%\OutlookAI\models.json        (per-user fallback)

Other users pick it up the next time Outlook starts; open task panes on
this machine refresh their effort lists right away. Until the first
refresh, and whenever a newer OutlookAI build ships a newer list than the
cached one, OutlookAI uses the list it shipped with. New OpenAI models and
effort levels appear after an Update Models; no reinstall needed.

When the catalog announces a retirement (e.g. gpt-5.5 on 2026-10-14),
Settings shows it and requests move to the announced replacement once the
date passes, even after the retired model drops out of the list (for up to
a year). The saved choice in config.xml is not rewritten; Settings opens on
the replacement, and picking a model there and saving makes it permanent. A
model that is merely
missing from the list (ChatGPT filters it by Codex version and plan) is not
treated as retired: if config names it, requests keep using it and Settings
says it isn't in the list.

The catalog request identifies as the latest Codex CLI release (looked up
on api.github.com/repos/openai/codex), because ChatGPT only lists models
that release supports. If that lookup is blocked, pin a version in
C:\Program Files\OutlookAI\config.xml:

  <ModelCatalogClientVersion>0.156.0</ModelCatalogClientVersion>


UPDATES
-------
Admins update OutlookAI from inside Outlook: Settings (gear icon, admin
password) -> Updates. Check Now looks up the latest release on GitHub;
Install Update, offered when that release is newer, downloads its zip,
checks the SHA256 and runs its installer with administrator rights (UAC
prompt). The installer closes Outlook for every user on the machine and
leaves it closed, so warn RDS users first; everyone reopens Outlook when
it finishes. Updates keep C:\Program Files\OutlookAI\config.xml and the
Settings in C:\ProgramData\OutlookAI.

The updater needs HTTPS access to api.github.com, github.com and the
*.githubusercontent.com hosts GitHub redirects release downloads to.
Downloads go to %LOCALAPPDATA%\OutlookAI\Updates\<tag>\ and each attempt
is logged in %LOCALAPPDATA%\OutlookAI\update-history.json.

To update without the updater, get the new install bundle and run its
installer as for a fresh install; it upgrades in place.


ROTATING CREDENTIALS
--------------------
OutlookAI can't revoke tokens remotely; rotation is a manual two-step:

1. On the RDS server, as any user who knows the OutlookAI admin password:
   - Open Outlook -> AI Assistant -> gear icon (Settings).
   - Enter the OutlookAI admin password.
   - Click "Sign Out" in the ChatGPT Account section.
   - Click "Sign In" and authenticate with the new ChatGPT account.
   - If the new account is on a different plan, click AI Behavior ->
     Update Models so the model list matches it (Settings flags a list
     fetched for a different account).

2. From the OpenAI side (recommended after any suspected leak):
   - Sign the previous account out at https://chatgpt.com/#settings
     (rotates session tokens server-side).


SHAPE C - IT-MANAGED IMAGE / SILENT INSTALL
--------------------------------------------
Use case: corporate Windows image, MDT/SCCM rollout, or any deployment
where the install must complete without operator interaction.

If the image is built without internet access, install the WebView2
runtime into it first (Microsoft's Evergreen Standalone Installer); the
bootstrapper in the bundle downloads the runtime.

Then bake the installer into the image:

1. Copy the extracted install bundle (see GET THE INSTALL BUNDLE) to a
   known location on the gold image, e.g. C:\OutlookAI.

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

Install or update fails, or OutlookAI.dll goes missing afterwards:
  - Antivirus may be quarantining the add-in's files. Exclude
    C:\Program Files\OutlookAI\ and %LOCALAPPDATA%\OutlookAI\Updates\
    (all scan engines, including behavior monitoring), then run the
    installer again.

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
2. Run Uninstall-OutlookAI.ps1 from the install bundle (or from Deploy\
   in a clone of the repository), e.g.:
     C:\OutlookAI\Uninstall-OutlookAI.ps1

This removes:
  - HKLM Outlook add-in registration (64-bit + WOW6432Node).
  - C:\Program Files\OutlookAI install directory.
  - C:\ProgramData\OutlookAI\auth.json + sidecar refresh lock.

It keeps, on purpose, so a reinstall picks them up:
  - C:\ProgramData\OutlookAI\config.xml (the Settings for every user,
    including the admin password) and models.json.
  - C:\ProgramData\OutlookAI\Backups, for rollback.
Delete C:\ProgramData\OutlookAI as well to remove everything. Each user's
%LOCALAPPDATA%\OutlookAI and %APPDATA%\OutlookAI stay in their profile.


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
