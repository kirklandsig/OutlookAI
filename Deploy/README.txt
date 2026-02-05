OutlookAI - RDS Deployment Guide
==================================

PREREQUISITES
-------------
- Windows Server 2022 or 2025
- .NET Framework 4.8
- Visual Studio Tools for Office Runtime
  Download: https://aka.ms/VSTORuntime


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
   .\Install-OutlookAI.ps1 -SourcePath "C:\OutlookAI"


VERIFICATION
------------
1. Log in as a user
2. Open Outlook
3. Click "New Email" to open compose window
4. Look for "AI Assistant" button in the ribbon
5. Or check: File > Options > Add-ins > OutlookAI should be listed


UNINSTALLATION
--------------
1. Open PowerShell as Administrator
2. Run: .\Uninstall-OutlookAI.ps1


TROUBLESHOOTING
---------------

Add-in shows in list but won't load / keeps unchecking:

1. Check if VSTO Runtime is installed:
   - Look for "Microsoft Visual Studio 2010 Tools for Office Runtime"
     in Programs and Features
   - If missing, download from: https://aka.ms/VSTORuntime

2. Check Outlook's disabled items:
   - File > Options > Add-ins
   - At bottom: "Manage" dropdown > select "Disabled Items" > Go
   - If OutlookAI is listed, select it and click "Enable"

3. Check for COM Add-in issues:
   - File > Options > Add-ins
   - At bottom: "Manage" dropdown > select "COM Add-ins" > Go
   - Check the box next to OutlookAI
   - If you get an error, note the message

4. Check Windows Event Viewer:
   - Open Event Viewer
   - Windows Logs > Application
   - Look for errors from "Outlook" or ".NET Runtime"

5. Verify files exist:
   - Check C:\Program Files\OutlookAI\ contains:
     * OutlookAI.dll
     * OutlookAI.vsto
     * OutlookAI.dll.manifest
     * NAudio.Core.dll
     * NAudio.WinMM.dll

6. Re-register with full trust (run as admin):

   $path = "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI"
   Set-ItemProperty $path -Name "LoadBehavior" -Value 3 -Type DWord

   $trust = "HKLM:\SOFTWARE\Microsoft\.NETFramework\Security\TrustManager\PromptingLevel"
   Set-ItemProperty $trust -Name "MyComputer" -Value "Enabled"
   Set-ItemProperty $trust -Name "LocalIntranet" -Value "Enabled"

7. Check for 32-bit vs 64-bit Office mismatch:
   - Find out if Office is 32 or 64 bit:
     Outlook > File > Office Account > About Outlook
   - If 32-bit, check registry exists at:
     HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins\OutlookAI


Add-in not showing at all:
- Make sure Outlook was restarted after install
- Verify registry key exists (see above)

Trust errors:
- The install script configures trust settings
- If issues persist, check Group Policy isn't blocking add-ins
- Group Policy location:
  User Config > Admin Templates > Microsoft Outlook > Security


FILES INSTALLED
---------------
Location: C:\Program Files\OutlookAI\
- OutlookAI.dll          (main add-in)
- OutlookAI.vsto         (deployment manifest)
- OutlookAI.dll.manifest (application manifest)
- NAudio*.dll            (audio recording for voice input)
- Microsoft.Office.Tools.*.dll (VSTO utilities)

Registry: HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins\OutlookAI


SUPPORT
-------
For issues, contact your IT administrator.
