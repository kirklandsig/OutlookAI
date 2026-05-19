# Installing OutlookAI

OutlookAI supports three install shapes. All three share the same installer
(`Deploy/Install-OutlookAI.ps1`) and the same OAuth flow (sign in once with
your ChatGPT account, then OutlookAI uses your existing subscription for
inference).

| Shape | Use case | Detail |
|---|---|---|
| **Single workstation** | One developer or power user. | [Deploy/README.txt — Shape A](../Deploy/README.txt) |
| **Multi-user RDS / Terminal Server** | Shared server, many interactive users, one shared ChatGPT credential. | [Deploy/README.txt — Shape B](../Deploy/README.txt) |
| **IT-managed image / silent install** | MDT / SCCM / corporate gold image. | [Deploy/README.txt — Shape C](../Deploy/README.txt) |

## Quick start (single workstation)

```powershell
git clone https://github.com/kirklandsig/OutlookAI.git
cd OutlookAI

# Publish Release into a staging folder
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "VSTO2\OutlookAI.sln" /target:Publish /p:Configuration=Release /p:Platform="Any CPU" `
  /p:PublishDir="C:\OutlookAI\"

# Install elevated
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
.\Deploy\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"

# Open Outlook → AI Assistant → sign in with your ChatGPT account.
```

For the full deployment story (cleanup, shared credentials, rotation,
troubleshooting, rollback, uninstall), see
[`Deploy/README.txt`](../Deploy/README.txt). That file is the canonical
install guide; this page is a pointer.
