# Installing OutlookAI

OutlookAI supports three install shapes. All three share the same install
bundle and installer (`Install-OutlookAI.ps1`) and the same OAuth flow (sign
in once with your ChatGPT account, then OutlookAI uses your existing
subscription for inference).

| Shape | Use case | Detail |
|---|---|---|
| **Single workstation** | One developer or power user. | [Deploy/README.txt — Shape A](../Deploy/README.txt) |
| **Multi-user RDS / Terminal Server** | Shared server, many interactive users, one shared ChatGPT credential. | [Deploy/README.txt — Shape B](../Deploy/README.txt) |
| **IT-managed image / silent install** | MDT / SCCM / corporate gold image. | [Deploy/README.txt — Shape C](../Deploy/README.txt) |

## Quick start (single workstation)

Download `OutlookAI-vX.Y.Z-RDS-Deploy.zip` and its `.sha256` from the
[latest release](https://github.com/kirklandsig/OutlookAI/releases/latest),
then in an elevated PowerShell in the download folder:

```powershell
# Check the download (must print True) and extract it
$zip = ".\OutlookAI-vX.Y.Z-RDS-Deploy.zip"
(Get-FileHash $zip -Algorithm SHA256).Hash -eq (Get-Content "$zip.sha256").Trim()
Unblock-File $zip
Expand-Archive $zip -DestinationPath C:\OutlookAI

# Install
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
C:\OutlookAI\Install-OutlookAI.ps1 -SourcePath C:\OutlookAI

# Open Outlook → AI Assistant → sign in with your ChatGPT account.
```

After that, admins install new versions from Settings → **Updates**
(**Check Now**, then **Install Update**). To build the bundle from source,
see Contributing in the [README](../README.md#contributing).

For the full deployment story (shared credentials, Settings for every user,
updates, rotation, troubleshooting, rollback, uninstall), see
[`Deploy/README.txt`](../Deploy/README.txt). That file is the canonical
install guide; this page is a pointer.
