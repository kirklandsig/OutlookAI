# In-App Updater — Design

Branch: `feature/in-app-updater`
Author session: 2026-05-20
Status: Proposed (pending user review)

## Why

Today every release means a manual round-trip: publish Release, copy the
deploy ZIP to each RDS server, run `Install-OutlookAI.ps1` elevated, verify
the SHA256. That gate keeps the admin in the loop, which is good — but every
small fix (the schema-steering tweaks in issues #3 / #4 are a good example)
forces a full redeploy that nobody schedules on Friday afternoon.

OutlookAI v2 is going to keep getting small improvements. We want an admin to
be able to click a button in Settings, accept a UAC prompt, and have the same
`Install-OutlookAI.ps1` machinery they already trust run against a freshly
downloaded GitHub Release ZIP. Same install logic, same registry writes,
same file unblocks — just sourced from the cloud instead of a flash drive.

## Goal

Ship a "Check for updates" / "Install Update" workflow in the Settings UI
that:

1. Reads the running install's version from a `version.json` baked into the
   deploy package.
2. Calls `api.github.com/repos/kirklandsig/OutlookAI/releases/latest` and
   compares the returned `tag_name` to the installed version using semver.
3. Downloads the release ZIP + sidecar `.sha256`, verifies the hash,
   extracts to `%LOCALAPPDATA%\OutlookAI\Updates\<tag>\staging\`.
4. After a loud "this will close Outlook for all users" confirmation, spawns
   an elevated PowerShell that runs the extracted `Install-OutlookAI.ps1
   -SourcePath <staging>`. The installer kills all `OUTLOOK.exe` processes
   (its existing behavior) and the add-in dies with it. The elevated
   PowerShell window stays open (`-NoExit`) so the admin sees the install
   transcript.
5. Admin reopens Outlook; new build is live.

Success = the redeploy round-trip drops from a multi-step manual process to
a button click + UAC accept + Outlook restart.

## Non-goals

- ClickOnce auto-update or `<deploymentProvider>` rewiring. The current
  elevated-installer model is the source of truth.
- Background polling / daily auto-checks. Manual `Check Now` only.
- Beta / pre-release channel. The updater calls `/releases/latest`, which
  by GitHub's contract returns only the most recent non-draft, non-prerelease.
- Scheduled-task / after-hours install scheduling.
- Auto-relaunching Outlook on the admin's session after install (admin
  reopens manually).
- Notifying non-admin users that an update is happening. They will see
  Outlook close; that is acceptable for now.
- Downgrade ("install older version"). Only forward updates.
- Multi-server fan-out. Admin still runs the update on each RDS server
  separately.
- A separate `Updater.exe` binary. The add-in is the updater; the
  `Install-OutlookAI.ps1` it spawns also replaces the add-in itself on next
  Outlook start.
- Telemetry beyond a local `update-history.json` log file.

## Audience and constraints

**Primary audience:** the OutlookAI admin (the same person who knows the
admin password and currently runs `Install-OutlookAI.ps1` manually). On the
firm's RDS deployment this is the IT lead.

**Constraints inherited from the v2 install model:**

- `OutlookAI.dll` is memory-mapped by every running `Outlook.exe`. Windows
  holds an exclusive lock on the file. The DLL cannot be overwritten until
  every Outlook process exits. The existing installer already force-closes
  all Outlook instances on the machine; the updater inherits that.
- The install directory is `C:\Program Files\OutlookAI\` — writes require
  Administrator. The updater must elevate via UAC. End users cannot trigger
  the install action (no UAC for them).
- Per-user state (`%APPDATA%\OutlookAI\config.xml`) is preserved by the
  installer's existing logic. The updater does not touch it.
- The shared OAuth credential at `C:\ProgramData\OutlookAI\auth.json` is
  preserved across reinstalls. The updater does not touch it.

## Versioning scheme

Source of truth = semver release tags (`v2.1.0`, `v2.2.0`, `v2.2.0-beta.1`
for future betas if ever needed).

`version.json` baked into the deploy ZIP at CI time:

```json
{
  "tag": "v2.1.0",
  "commit": "abc1234",
  "build_date": "2026-06-02T19:14:00Z",
  "repo": "kirklandsig/OutlookAI"
}
```

Installed location after `Install-OutlookAI.ps1` runs: `C:\Program
Files\OutlookAI\version.json`. The installer gets one new step that copies
this file from the staging dir into the install dir.

If `version.json` is absent (e.g. a pre-this-feature build), the add-in
treats the install as a "dev build" and shows `Current: (unknown)` in the
UI; updates are still allowed.

## Architecture

Five small components, each with one clear responsibility:

```
SettingsForm (Updates section)
    │
    │  Check Now
    │  ──────────────► GitHubReleaseClient.GetLatestStableAsync
    │                    │  GET https://api.github.com/repos/.../releases/latest
    │                    ▼
    │                  ReleaseInfo { Tag, ZipUrl, ShaUrl, ReleasePageUrl, PublishedAt }
    │                    │
    │                    ▼
    │                  VersionComparator.Compare(installedTag, latestTag)
    │                    │
    │                    ▼
    │                  UpdateState { Availability, ReleaseInfo }
    │
    │  Install Update
    │  ──────────────► MessageBox confirm (loud RDS warning)
    │                    │
    │                    ▼
    │                  UpdateDownloader.DownloadAsync
    │                    │  HTTP GET .zip, GET .sha256
    │                    │  verify hash
    │                    │  extract to %LOCALAPPDATA%\OutlookAI\Updates\<tag>\staging\
    │                    ▼
    │                  DownloadedUpdate { StagingDir, InstallerScriptPath }
    │                    │
    │                    ▼
    │                  UpdateInstaller.LaunchElevatedInstall
    │                    │  Process.Start(powershell.exe -NoExit ... runas)
    │                    ▼
    │                  Detached elevated PowerShell
    │                    │  Install-OutlookAI.ps1 kills OUTLOOK.exe
    │                    │  add-in dies; installer runs to completion
    │                    │  PS window stays open with transcript
    │                    │
    │                    ▼
    │                  Admin closes window, reopens Outlook
    │                    │
    │                    ▼
    │                  New version.json reads v2.1.0 next time Settings opens
```

### Components

**`Services/Updates/UpdateManifest.cs`** — DTO + parser.

- `string Tag` (e.g. `"v2.1.0"`).
- `string Commit` (short SHA, optional).
- `DateTimeOffset BuildDate`.
- `string Repo`.
- `static UpdateManifest LoadFromInstallDir()` — reads
  `C:\Program Files\OutlookAI\version.json`. Returns a `DevBuild` sentinel
  manifest (`Tag = "(dev build)"`) if the file is missing or malformed.
- `static UpdateManifest LoadFromZip(string zipPath)` — opens the ZIP
  without full extraction, reads the inner `version.json`, returns the same
  shape.

**`Services/Updates/GitHubReleaseClient.cs`** — async `HttpClient` wrapper.

- Constructor takes an `HttpClient` (test-injectable; existing
  `FakeHttpMessageHandler` helper covers it).
- `Task<ReleaseLookupResult> GetLatestStableAsync(CancellationToken ct)` —
  `GET https://api.github.com/repos/kirklandsig/OutlookAI/releases/latest`
  with `User-Agent: OutlookAI-Updater/<currentTag>` (falling back to
  `OutlookAI-Updater/dev` if `UpdateManifest.LoadFromInstallDir().Tag` is
  the dev sentinel) and `Accept: application/vnd.github+json` headers.
  Returns one of:
  - `ReleaseFound { ReleaseInfo info }`
  - `NoReleasesAvailable` (on HTTP 404 from `/releases/latest`)
  - `RateLimited { DateTimeOffset resetAt }` (on HTTP 403 with
    `X-RateLimit-Remaining: 0`)
  - `NetworkError { string detail }` (on transport failures)
- Default `HttpClient` honors the system proxy
  (`WebRequest.DefaultWebProxy`), so Zscaler / corporate forced proxies just
  work without extra config.

**`Services/Updates/ReleaseInfo.cs`** — DTO.

- `string Tag` (`tag_name`).
- `string ReleasePageUrl` (`html_url`).
- `DateTimeOffset PublishedAt` (`published_at`).
- `string Body` (release notes; used for UI tooltip).
- `string ZipAssetName` (e.g. `OutlookAI-v2.1.0-RDS-Deploy.zip`).
- `string ZipUrl` (browser download URL).
- `string ShaUrl` (browser download URL for the `.sha256` sidecar; null if
  the asset is missing).

**`Services/Updates/VersionComparator.cs`** — pure static helper.

- `static UpdateAvailability Compare(string installedTag, string latestTag)`.
- Strips `v` / `V` prefix.
- Parses `major.minor.patch[-prerelease[.N]]`. Prerelease suffix sorts
  *lower* than the same numeric tuple without it (so `2.1.0-beta.1 <
  2.1.0`). Compares left-to-right by numeric component; falls back to
  lexicographic on the prerelease suffix.
- Returns:
  - `NoUpdate` (installed >= latest)
  - `NewerAvailable` (latest > installed)
  - `NotComparable` (either tag fails to parse; e.g. `(dev build)`)
  - `NoReleases` (delegated through from the client)

**`Services/Updates/UpdateDownloader.cs`** — file work.

- Constructor takes an `HttpClient`, a logger, and an `IFileSystem` thin
  wrapper (so tests can inject an in-memory FS without touching
  `%LOCALAPPDATA%`).
- `Task<DownloadResult> DownloadAsync(ReleaseInfo info, IProgress<int>
  progress, CancellationToken ct)`. Steps:
  1. Compute staging path: `%LOCALAPPDATA%\OutlookAI\Updates\<tag>\`.
     Delete any prior contents at this path (idempotent retry).
  2. HTTP GET the `.zip` and `.sha256` URLs in parallel; stream the ZIP
     to disk with progress reporting (chunk size 64 KB).
  3. Compute SHA256 of the downloaded ZIP; compare to the sidecar's
     contents (sidecar is a single line of lowercase hex).
  4. Extract ZIP to `<staging>\extracted\`. Verify
     `<staging>\extracted\Install-OutlookAI.ps1` exists.
  5. Return `DownloadResult.Success { StagingDir, InstallerScriptPath,
     ExtractedDir, ExpectedSha256 }`.
- Returns typed failure cases:
  - `HashMismatch { Expected, Actual }`
  - `DownloadFailed { string Detail }`
  - `MissingInstallerScript`
  - `Cancelled`
- After a successful install (detected at next Outlook startup), cleans up
  older `Updates\<tag>\` siblings, keeping the most recent + the currently
  installed tag for diagnostics.

**`Services/Updates/UpdateInstaller.cs`** — launches the elevated install.

- `LaunchResult LaunchElevatedInstall(DownloadResult.Success update)`.
- Builds the elevated `Process.Start`:

```csharp
var psi = new ProcessStartInfo("powershell.exe")
{
    UseShellExecute = true,
    Verb = "runas",
    WorkingDirectory = Path.GetTempPath(),
    Arguments = string.Format(
        "-NoExit -NoProfile -ExecutionPolicy Bypass -File \"{0}\" -SourcePath \"{1}\"",
        update.InstallerScriptPath,
        update.ExtractedDir),
};
```

- Returns:
  - `LaunchResult.Launched { int Pid }` on UAC accept.
  - `LaunchResult.UacDeclined` on `Win32Exception { NativeErrorCode == 1223 }`.
  - `LaunchResult.LaunchFailed { string Detail }` for anything else.
- Does **not** wait for the child process. The child is detached and will
  outlive Outlook getting killed by the installer.
- Writes a "launched" entry to the history log immediately so the next
  Outlook start can correlate.

**`Services/Updates/UpdateHistoryLog.cs`** — append-only structured log at
`%LOCALAPPDATA%\OutlookAI\update-history.json`. Keeps the last 50 entries;
oldest dropped. Schema:

```json
[
  { "ts": "2026-06-02T19:14:00Z", "action": "check",    "result": "newer_available", "tag": "v2.1.0", "details": "" },
  { "ts": "2026-06-02T19:14:30Z", "action": "download", "result": "ok",               "tag": "v2.1.0", "details": "sha256_ok" },
  { "ts": "2026-06-02T19:14:32Z", "action": "launch",   "result": "launched",         "tag": "v2.1.0", "details": "pid=12345" }
]
```

The Updates section of Settings shows the most recent successful install
(read from this log) as `Last installed: v2.1.0 (2026-06-02 19:14)`.

### UI (`SettingsForm.cs` and `.Designer.cs`)

A new `GroupBox` titled **Updates** below the existing AI Behavior section.
Layout (vertical):

- `Current: v2.0.0` (label, reads `UpdateManifest.LoadFromInstallDir().Tag`)
- `Latest: —` (label, populated after a check)
- `Last checked: —` (label, populated after a check)
- `[ Check Now ]` (button)
- `[ Install Update ]` (button, disabled by default)
- Status line (label, single line, shows progress / errors)

State machine in the form's code-behind:

- `Idle` → `Check Now` triggers `Checking`.
- `Checking` → on result, transitions to `UpdateAvailable` (enables Install
  Update), `UpToDate`, `NoReleases`, or `CheckFailed`.
- `UpdateAvailable` → `Install Update` shows the confirmation `MessageBox`;
  on OK, transitions to `Downloading`.
- `Downloading` → on success, `Launching`; on failure, `DownloadFailed`.
- `Launching` → on `Launched`, the form stays alive but adds a final status
  line: "Installer launched. Outlook will close shortly to apply the
  update."; the add-in process is going to die any second.
- `UacDeclined` / `LaunchFailed` keep state at `UpdateAvailable` so the
  admin can retry.

The big confirmation `MessageBox` copy:

```
⚠ Install OutlookAI v2.1.0

This will:
  • close Outlook for ALL users currently on this server
  • run the OutlookAI installer with administrator privileges
  • leave Outlook closed when finished — everyone reopens manually

Have you given users a heads-up?

      [ Cancel ]    [ Install Now ]
```

Captions and copy live in resource strings so the source-level test can pin
them.

## Release pipeline (`.github/workflows/release.yml`)

Trigger: push of a tag matching `v*` to `master`.

Steps:

1. Checkout at the tag.
2. Setup MSBuild + nuget (Windows runner: `windows-2025`).
3. `nuget restore VSTO2/OutlookAI.sln`.
4. MSBuild Release publish into `out/staging/`.
5. Generate `out/staging/version.json` with:
   - `tag` = `$GITHUB_REF_NAME`
   - `commit` = `$(git rev-parse --short HEAD)`
   - `build_date` = ISO 8601 UTC `now`
   - `repo` = `$GITHUB_REPOSITORY`
6. Copy install assets into `out/staging/` (same set we package today:
   `Install-OutlookAI.ps1`, `Uninstall-OutlookAI.ps1`,
   `Fetch-WebView2Bootstrapper.ps1`, `MicrosoftEdgeWebView2Setup.exe`,
   `Deploy/README.txt`, and `OutlookAI.dll` for hash convenience).
7. ZIP `out/staging/` into `out/OutlookAI-${{ github.ref_name
   }}-RDS-Deploy.zip`.
8. Compute SHA256 of the ZIP into `out/OutlookAI-${{ github.ref_name
   }}-RDS-Deploy.zip.sha256` (single line, lowercase hex, no filename
   suffix).
9. `gh release create` with both files as assets, body sourced from the
   annotated tag's message (`git tag -l --format='%(contents)' <tag>`).

Manual fallback: if Actions is unavailable, the same steps can run locally
(we already do steps 1–7 by hand; step 8 is `Get-FileHash`; step 9 is `gh
release create`).

## Failure modes (compact reference)

| Failure | UI / behavior |
|---|---|
| Network down during Check Now | "Could not reach GitHub. Check your network and try again." No state change. |
| `404` from `/releases/latest` | "No releases published yet on GitHub." `Install Update` disabled. |
| Unparseable tag (e.g. dev build) | `Latest: <tag>` with "(could not compare to installed)". `Install Update` disabled. |
| `Latest <= installed` | `Latest: v2.0.0 (already installed)`. `Install Update` disabled. |
| Hash mismatch | "Downloaded file failed integrity check. Aborting install." Staging cleaned up. |
| Download interrupted | "Download failed. Try again." Staging cleaned up. |
| Missing installer in ZIP | "Update package is malformed (no installer). Please file a bug." |
| UAC declined | "Update cancelled — administrator privileges required." Staging kept; admin can retry. |
| `Install-OutlookAI.ps1` exits non-zero | Can't detect from the add-in because it's detached and Outlook died. The PS window shows the transcript; admin reads it. Next Outlook start, the Updates section will either show the new version (install succeeded) or the old version (install failed). |
| Concurrent install attempt | We write `%LOCALAPPDATA%\OutlookAI\Updates\.in-progress` at launch time. On next successful Outlook start, `ThisAddIn.Startup` reads the sentinel; if the installed `version.json` now matches the sentinel's tag, the install succeeded — clear the sentinel. If the tag still doesn't match and the sentinel is older than 30 min, clear it (assume aborted) and log to history. If admin clicks Install Update while the sentinel is set and < 30 min old, show "An install appears to be in progress." Defensive. |
| `version.json` missing in install dir | Treat as dev build; allow updates anyway. |

## Honesty constraints

- The updater does **not** make installs safer or less disruptive than today
  — Outlook still closes for every user. The win is convenience, not
  reduced impact.
- The updater does **not** roll back automatically on failure. The existing
  installer's backup behavior (`C:\ProgramData\OutlookAI\Backups\`) is the
  only rollback path, same as today.
- The updater does **not** auto-update itself. If a future release changes
  the updater's API contract with GitHub, the old updater can still
  download the new ZIP and run its installer — the installer is what
  replaces the add-in, so the new updater is live on next Outlook start.
- The updater does **not** poll. The admin clicks Check Now. We don't
  surprise the firm with background traffic to `api.github.com`.

## Acceptance criteria

- A `git tag v<x.y.z> && git push --tags` push triggers `.github/workflows/release.yml`,
  which produces a GitHub Release with two assets attached:
  `OutlookAI-<tag>-RDS-Deploy.zip` and `OutlookAI-<tag>-RDS-Deploy.zip.sha256`.
- The deploy ZIP contains a `version.json` whose `tag` matches the pushed
  tag.
- A fresh install of the new ZIP via `Install-OutlookAI.ps1` writes
  `version.json` into `C:\Program Files\OutlookAI\`.
- After upgrading, Settings → Updates shows `Current: v<tag>` matching what
  was installed.
- Clicking Check Now against a repo with no releases yet shows "No releases
  published yet on GitHub" and leaves Install Update disabled.
- Clicking Check Now against a repo with a newer release enables Install
  Update.
- Clicking Install Update shows the confirmation dialog with the RDS
  warning copy.
- Confirming triggers download → hash verify → UAC prompt → detached
  elevated PowerShell window.
- Tampering with the ZIP locally (e.g. flipping a byte) causes the
  downloader to fail with `HashMismatch` and clean up the staging dir.
- Tests: `Total tests: 553 + ~20 new` (estimate ~20 across the five new
  components and the source-level UI test).
- Build: VS MSBuild Debug Any CPU green with only the existing `MSB3277`
  warning.
- No regressions in existing 553 tests.

## Out of scope for this pass (explicit follow-ups)

- Background update polling and a passive "Update available" badge in chat
  (Option C from the earlier discussion).
- Beta channel toggle.
- Scheduled-task / after-hours installs.
- Auto-relaunch Outlook for the admin's session post-install.
- Notifying non-admin users that an update is happening.
- Signed release artifacts (Authenticode on the DLL, signed installer
  script). The SHA256 sidecar is the current integrity story.
- Auto-updating across multiple RDS servers from a single admin click.
