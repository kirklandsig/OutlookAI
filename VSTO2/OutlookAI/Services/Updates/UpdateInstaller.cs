using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;

namespace OutlookAI.Services.Updates
{
    /// <summary>
    /// Launches the extracted Install-OutlookAI.ps1 elevated via UAC. The
    /// process is detached: this call returns once the user accepts UAC, and
    /// the elevated PowerShell window will outlive Outlook getting killed by
    /// the installer.
    /// </summary>
    public sealed class UpdateInstaller
    {
        public LaunchResult LaunchElevatedInstall(DownloadSuccess update)
        {
            if (update == null) throw new ArgumentNullException(nameof(update));
            if (string.IsNullOrWhiteSpace(update.InstallerScriptPath)) throw new ArgumentException("InstallerScriptPath required.", nameof(update));
            if (string.IsNullOrWhiteSpace(update.ExtractedDir)) throw new ArgumentException("ExtractedDir required.", nameof(update));

            var psi = new ProcessStartInfo("powershell.exe")
            {
                UseShellExecute = true,
                Verb = "runas",
                // Force a local working directory so the elevated process does
                // not inherit a UNC CWD (Folder-Redirected Documents).
                WorkingDirectory = Path.GetTempPath(),
                Arguments = string.Format(
                    "-NoExit -NoProfile -ExecutionPolicy Bypass -File \"{0}\" -SourcePath \"{1}\"",
                    update.InstallerScriptPath,
                    update.ExtractedDir),
            };

            try
            {
                using (var p = Process.Start(psi))
                {
                    if (p == null)
                        return new LaunchFailed { Detail = "Process.Start returned null; elevated installer did not launch." };
                    return new Launched { Pid = p.Id };
                }
            }
            catch (Win32Exception ex) when (ex.NativeErrorCode == 1223)
            {
                // ERROR_CANCELLED — user clicked No on the UAC prompt.
                return new UacDeclined();
            }
            catch (Exception ex)
            {
                return new LaunchFailed { Detail = ex.Message };
            }
        }
    }
}
