using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.Deploy
{
    /// <summary>
    /// Runs parts of Deploy/Install-OutlookAI.ps1 in Windows PowerShell: its
    /// config helper functions, and whole install steps against temp folders.
    /// The rest of the installer never executes.
    /// </summary>
    public class InstallScriptConfigTests
    {
        private const string AuthPath = @"C:\ProgramData\OutlookAI\auth.json";

        private const string EarlyV2Global =
            @"<Config><AdminPassword>a</AdminPassword><CodexAuthPath>C:\ProgramData\OutlookAI\auth.json</CodexAuthPath>"
            + "<Model>gpt-5.5</Model><VoiceModel>gpt-realtime-1.5</VoiceModel><MaxTokens>65536</MaxTokens></Config>";

        private static readonly Encoding Windows1252 = Encoding.GetEncoding(1252);

        private static string PsQuote(string value) => "'" + value.Replace("'", "''") + "'";

        // Runs `body` after loading the installer's *OutlookAI*Config* functions and
        // returns its "key=value" output lines. In `body`, `Get-InstallerStep 5`
        // returns the installer's step 5 as a script block.
        private static Dictionary<string, string> RunInstallerHelpers(string body)
        {
            var installer = PsQuote(RepoFiles.Find("Deploy", "Install-OutlookAI.ps1"));
            var harness = Path.Combine(Path.GetTempPath(), "outlookai-install-test-" + Guid.NewGuid().ToString("N") + ".ps1");
            var script = new StringBuilder()
                .AppendLine("$ErrorActionPreference = 'Stop'")
                .AppendLine("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8")
                .AppendLine("$tokens = $null; $errors = $null")
                .AppendLine("$ast = [System.Management.Automation.Language.Parser]::ParseFile(" + installer + ", [ref]$tokens, [ref]$errors)")
                .AppendLine("if ($errors.Count -gt 0) { throw ('installer parse errors: ' + $errors.Count) }")
                .AppendLine("$functions = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -like '*OutlookAI*Config*' }, $true)")
                .AppendLine("if (@($functions).Count -lt 2) { throw 'installer config helpers not found' }")
                .AppendLine("foreach ($f in $functions) { . ([ScriptBlock]::Create($f.Extent.Text)) }")
                .AppendLine("$installerText = [System.IO.File]::ReadAllText(" + installer + ")")
                .AppendLine("function Get-InstallerStep([int]$Step) {")
                .AppendLine("    $start = $installerText.IndexOf('# --- ' + $Step + '.')")
                .AppendLine("    $end = $installerText.IndexOf('# --- ' + ($Step + 1) + '.')")
                .AppendLine("    if ($start -lt 0 -or $end -le $start) { throw ('installer step ' + $Step + ' not found') }")
                .AppendLine("    [ScriptBlock]::Create($installerText.Substring($start, $end - $start))")
                .AppendLine("}")
                .AppendLine(body)
                .ToString();
            File.WriteAllText(harness, script, new UTF8Encoding(true));
            try
            {
                var psi = new ProcessStartInfo(
                    "powershell.exe",
                    "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + harness + "\"")
                {
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                };
                using (var process = Process.Start(psi))
                {
                    var stdout = process.StandardOutput.ReadToEndAsync();
                    var stderr = process.StandardError.ReadToEndAsync();
                    if (!process.WaitForExit(60000))
                    {
                        try { process.Kill(); } catch { }
                        Assert.True(false, "PowerShell harness timed out");
                    }
                    Assert.True(process.ExitCode == 0, "PowerShell harness failed: " + stderr.Result + stdout.Result);
                    var values = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var line in stdout.Result.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var eq = line.IndexOf('=');
                        if (eq > 0) values[line.Substring(0, eq)] = line.Substring(eq + 1);
                    }
                    return values;
                }
            }
            finally
            {
                try { File.Delete(harness); } catch { }
            }
        }

        // Temp folders standing in for C:\Program Files\OutlookAI (with leftovers
        // from the previous version), the Backups folder, the release package
        // and C:\Users, for running installer steps.
        private sealed class InstallSandbox : IDisposable
        {
            public InstallSandbox()
            {
                Root = Path.Combine(Path.GetTempPath(), "outlookai-install-steps-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(Path.Combine(InstallPath, "Application Files", "OutlookAI_1_0_0_4"));
                File.WriteAllText(Path.Combine(InstallPath, "Application Files", "OutlookAI_1_0_0_4", "OutlookAI.dll.deploy"), "old");
                File.WriteAllText(Path.Combine(InstallPath, "OutlookAI.dll"), "old");
                Directory.CreateDirectory(BackupRoot);
                Directory.CreateDirectory(SourcePath);
                Directory.CreateDirectory(UsersRoot);
            }

            public string Root { get; }
            public string InstallPath => Path.Combine(Root, "OutlookAI");
            public string ConfigPath => Path.Combine(InstallPath, "config.xml");
            public string BackupRoot => Path.Combine(Root, "Backups");
            public string SourcePath => Path.Combine(Root, "Source");
            public string UsersRoot => Path.Combine(Root, "Users");
            public string ProgramDataPath => Path.Combine(Root, "ProgramData");

            public string UserConfigPath(string user) => Path.Combine(UsersRoot, user, "AppData", "Roaming", "OutlookAI", "config.xml");

            // A backup keeps the write time of the config.xml it copied.
            public void WriteBackup(string nameTimestamp, DateTime writtenUtc, string content)
            {
                var path = Path.Combine(BackupRoot, "config.xml.v1.backup." + nameTimestamp);
                File.WriteAllText(path, content);
                File.SetLastWriteTimeUtc(path, writtenUtc);
            }

            // Each item is an installer step number, or PowerShell to run at that point.
            public Dictionary<string, string> Run(params string[] stepsAndCode)
            {
                var body = new StringBuilder()
                    .AppendLine("$InstallPath = " + PsQuote(InstallPath))
                    .AppendLine("$ConfigFilePath = Join-Path $InstallPath 'config.xml'")
                    .AppendLine("$BackupRoot = " + PsQuote(BackupRoot))
                    .AppendLine("$SourcePath = " + PsQuote(SourcePath))
                    .AppendLine("$UsersRoot = " + PsQuote(UsersRoot))
                    .AppendLine("$ProgramDataPath = " + PsQuote(ProgramDataPath))
                    .AppendLine("$AuthFilePath = " + PsQuote(AuthPath))
                    .AppendLine("$Timestamp = '20260922-120000'");
                foreach (var item in stepsAndCode)
                {
                    int step;
                    body.AppendLine(int.TryParse(item, out step) ? ". (Get-InstallerStep " + step + ") | Out-Null" : item);
                }
                body.AppendLine("'left=' + ((Get-ChildItem -LiteralPath $InstallPath -Force | ForEach-Object { $_.Name } | Sort-Object) -join ',')");
                return RunInstallerHelpers(body.ToString());
            }

            public void Dispose()
            {
                try
                {
                    foreach (var file in Directory.GetFiles(Root, "*", SearchOption.AllDirectories))
                    {
                        File.SetAttributes(file, FileAttributes.Normal);
                    }
                    Directory.Delete(Root, true);
                }
                catch { }
            }
        }

        private static string[] ElementNames(string path) =>
            XDocument.Load(path).Root.Elements().Select(e => e.Name.LocalName).ToArray();

        [Fact]
        public void V1Detection_MatchesOnlyClaudeEraFiles()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "'v1-keys=' + (Test-OutlookAIV1Config ([xml]'<Config><ApiKey>sk-ant</ApiKey><AdminPassword>a</AdminPassword><Model>claude-3</Model><MaxTokens>4096</MaxTokens></Config>'))",
                "'v1-whisper=' + (Test-OutlookAIV1Config ([xml]'<Config><WhisperModel>whisper-1</WhisperModel></Config>'))",
                "'v1-claude-model=' + (Test-OutlookAIV1Config ([xml]'<Config><AdminPassword>a</AdminPassword><Model> Claude-Opus-4-6 </Model></Config>'))",
                // What v2's Settings -> Save writes, and what early v2 builds wrote.
                "'v2-settings=' + (Test-OutlookAIV1Config ([xml]'<Config><AdminPassword>a</AdminPassword><Model>gpt-5.5</Model><ReasoningEffort>Medium</ReasoningEffort><WriteToolsEnabled>true</WriteToolsEnabled><EnabledWriteTools>outlook_create_draft</EnabledWriteTools></Config>'))",
                "'v2-password-only=' + (Test-OutlookAIV1Config ([xml]'<Config><AdminPassword>a</AdminPassword></Config>'))",
                "'v2-early-global=' + (Test-OutlookAIV1Config ([xml]'" + EarlyV2Global + "'))",
                // Early v2 with CodexAuthPath removed by hand: MaxTokens alone isn't v1.
                "'v2-early-no-authpath=' + (Test-OutlookAIV1Config ([xml]'<Config><AdminPassword>a</AdminPassword><Model>gpt-5.5</Model><MaxTokens>65536</MaxTokens><MaxBulkExportRows>500</MaxBulkExportRows></Config>'))",
                "'maxtokens-only=' + (Test-OutlookAIV1Config ([xml]'<Config><AdminPassword>a</AdminPassword><MaxTokens>65536</MaxTokens></Config>'))",
                // v1 leftovers next to v2 settings: kept as v2.
                "'mixed=' + (Test-OutlookAIV1Config ([xml]'<Config><ApiKey>x</ApiKey><Model>claude-3</Model><VoiceModel>gpt-realtime-2</VoiceModel></Config>'))",
                "'empty=' + (Test-OutlookAIV1Config ([xml]'<Config/>'))",
                "'not-config=' + (Test-OutlookAIV1Config ([xml]'<Other><ApiKey>x</ApiKey></Other>'))"));

            Assert.Equal("True", r["v1-keys"]);
            Assert.Equal("True", r["v1-whisper"]);
            Assert.Equal("True", r["v1-claude-model"]);
            Assert.Equal("False", r["v2-settings"]);
            Assert.Equal("False", r["v2-password-only"]);
            Assert.Equal("False", r["v2-early-global"]);
            Assert.Equal("False", r["v2-early-no-authpath"]);
            Assert.Equal("False", r["maxtokens-only"]);
            Assert.Equal("False", r["mixed"]);
            Assert.Equal("False", r["empty"]);
            Assert.Equal("False", r["not-config"]);
        }

        [Fact]
        public void GlobalConfig_AV2FileWithAnAuthPath_IsKeptAsIs()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "$previous = [Text.Encoding]::UTF8.GetBytes(\"<Config>`r`n    <AdminPassword>s3cret</AdminPassword>`r`n    <CodexAuthPath>D:\\auth\\auth.json</CodexAuthPath>`r`n    <VoiceModel>gpt-realtime-2</VoiceModel>`r`n    <MaxBulkExportRows>500</MaxBulkExportRows>`r`n</Config>`r`n\")",
                "'keep=' + ($null -eq (Get-OutlookAIGlobalConfig -PreviousBytes $previous))",
                "'early-keep=' + ($null -eq (Get-OutlookAIGlobalConfig -PreviousBytes ([Text.Encoding]::UTF8.GetBytes('" + EarlyV2Global + "'))))"));

            Assert.Equal("True", r["keep"]);
            Assert.Equal("True", r["early-keep"]);
        }

        [Fact]
        public void GlobalConfig_AddsAMissingAuthPath_AndKeepsTheRest()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "$previous = [Text.Encoding]::UTF8.GetBytes('<Config><AdminPassword>s3cret</AdminPassword><Model>gpt-5.5</Model><VoiceModel>gpt-realtime-2</VoiceModel><MaxTokens>65536</MaxTokens><MaxBulkExportRows>750</MaxBulkExportRows></Config>')",
                "$out = [xml](Get-OutlookAIGlobalConfig -PreviousBytes $previous)",
                "foreach ($name in 'CodexAuthPath', 'AdminPassword', 'Model', 'VoiceModel', 'MaxBulkExportRows') { $name + '=' + $out.DocumentElement.SelectSingleNode($name).InnerText }"));

            Assert.Equal(AuthPath, r["CodexAuthPath"]);
            Assert.Equal("s3cret", r["AdminPassword"]);
            Assert.Equal("gpt-5.5", r["Model"]);
            Assert.Equal("gpt-realtime-2", r["VoiceModel"]);
            Assert.Equal("750", r["MaxBulkExportRows"]);
        }

        [Fact]
        public void GlobalConfig_ReadsTheFileInItsDeclaredEncoding()
        {
            // "pässwörd" in a windows-1252 file that says so, as the add-in's XDocument.Load reads it.
            var r = RunInstallerHelpers(string.Join("\n",
                "$password = 'p' + [char]0xE4 + 'ssw' + [char]0xF6 + 'rd'",
                "$enc = [Text.Encoding]::GetEncoding(1252)",
                "$v2 = $enc.GetBytes('<?xml version=\"1.0\" encoding=\"windows-1252\"?><Config><AdminPassword>' + $password + '</AdminPassword></Config>')",
                "$v1 = $enc.GetBytes('<?xml version=\"1.0\" encoding=\"windows-1252\"?><Config><ApiKey>x</ApiKey><AdminPassword>' + $password + '</AdminPassword></Config>')",
                "'v2=' + (([xml](Get-OutlookAIGlobalConfig -PreviousBytes $v2)).DocumentElement.SelectSingleNode('AdminPassword').InnerText -ceq $password)",
                "'v1=' + (([xml](Get-OutlookAIGlobalConfig -PreviousBytes $v1)).DocumentElement.SelectSingleNode('AdminPassword').InnerText -ceq $password)"));

            Assert.Equal("True", r["v2"]);
            Assert.Equal("True", r["v1"]);
        }

        [Fact]
        public void GlobalConfig_V1EmptyAndMissingFiles_GetTheTemplate()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "function Describe([string]$text) { $x = [xml]$text; ($x.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' } | ForEach-Object { $_.Name + ':' + $_.InnerText }) -join '|' }",
                "function Bytes([string]$text) { [Text.Encoding]::UTF8.GetBytes($text) }",
                "'v1=' + (Describe (Get-OutlookAIGlobalConfig -PreviousBytes (Bytes '<Config><ApiKey>sk-ant</ApiKey><AdminPassword>old&amp;pw</AdminPassword><Model>claude-opus-4-6</Model><MaxTokens>4096</MaxTokens><WhisperModel>whisper-1</WhisperModel></Config>')))",
                "'v1-blank-password=' + (Describe (Get-OutlookAIGlobalConfig -PreviousBytes (Bytes '<Config><ApiKey>x</ApiKey><AdminPassword> </AdminPassword></Config>')))",
                "'empty-file=' + (Describe (Get-OutlookAIGlobalConfig -PreviousBytes ([byte[]]@())))",
                "'no-file=' + (Describe (Get-OutlookAIGlobalConfig -PreviousBytes $null -BackupBytes $null))"));

            // Only the AdminPassword comes from a v1 file: no v1 keys or Claude model.
            Assert.Equal(@"AdminPassword:old&pw|CodexAuthPath:" + AuthPath, r["v1"]);
            Assert.Equal(@"AdminPassword:admin|CodexAuthPath:" + AuthPath, r["v1-blank-password"]);
            Assert.Equal(@"AdminPassword:admin|CodexAuthPath:" + AuthPath, r["empty-file"]);
            Assert.Equal(@"AdminPassword:admin|CodexAuthPath:" + AuthPath, r["no-file"]);
        }

        [Fact]
        public void GlobalConfig_WithoutAConfig_TakesWhatEarlierInstallersCarriedOverFromTheBackup()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "function Describe([string]$text) { $x = [xml]$text; ($x.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq 'Element' } | ForEach-Object { $_.Name + ':' + $_.InnerText }) -join '|' }",
                "function FromBackup([string]$text) { Describe (Get-OutlookAIGlobalConfig -PreviousBytes $null -BackupBytes ([Text.Encoding]::UTF8.GetBytes($text))) }",
                "'v2=' + (FromBackup '<Config><AdminPassword>s3cret</AdminPassword><CodexAuthPath>D:\\x\\auth.json</CodexAuthPath><Model>gpt-6-sol</Model><VoiceModel>gpt-realtime-2</VoiceModel><MaxBulkExportRows>500</MaxBulkExportRows><ModelCatalogClientVersion>0.160.0</ModelCatalogClientVersion></Config>')",
                "'v1=' + (FromBackup '<Config><ApiKey>x</ApiKey><AdminPassword>v1pw</AdminPassword><Model>claude-opus-4-6</Model></Config>')",
                "'odd-shapes=' + (FromBackup '<Config><AdminPassword>from&lt;backup&gt;&amp;</AdminPassword><Model>bad model!</Model><ModelCatalogClientVersion>latest</ModelCatalogClientVersion><VoiceModel>x</VoiceModel></Config>')",
                "'not-config=' + (FromBackup '<Settings><AdminPassword>s3cret</AdminPassword><Model>gpt-6-sol</Model></Settings>')",
                "'not-xml=' + (FromBackup '<Config><AdminPassword>s3c')"));

            // Backups is writable by any signed-in user, so only these three fields come from it.
            Assert.Equal(@"AdminPassword:s3cret|CodexAuthPath:" + AuthPath + "|Model:gpt-6-sol|ModelCatalogClientVersion:0.160.0", r["v2"]);
            Assert.Equal(@"AdminPassword:v1pw|CodexAuthPath:" + AuthPath, r["v1"]);
            Assert.Equal(@"AdminPassword:from<backup>&|CodexAuthPath:" + AuthPath, r["odd-shapes"]);
            Assert.Equal(@"AdminPassword:admin|CodexAuthPath:" + AuthPath, r["not-config"]);
            Assert.Equal(@"AdminPassword:admin|CodexAuthPath:" + AuthPath, r["not-xml"]);
        }

        [Fact]
        public void GlobalConfig_AFileThatIsNotValidXml_IsLeftForTheAdminToFix()
        {
            var r = RunInstallerHelpers(
                "'keep=' + ($null -eq (Get-OutlookAIGlobalConfig -PreviousBytes ([Text.Encoding]::UTF8.GetBytes('<Config><AdminPassword>s3cret</AdminPassword><Model>gpt-5.5</Config>')) 3>$null))");

            Assert.Equal("True", r["keep"]);
        }

        [Fact]
        public void Update_KeepsAV2ConfigByteForByte_AndClearsTheRestOfTheInstallFolder()
        {
            using (var sandbox = new InstallSandbox())
            {
                // Hand-edited: no BOM, 4-space indent, a non-ASCII password.
                var original = new UTF8Encoding(false).GetBytes(
                    "<Config>\r\n    <AdminPassword>p\u00E9ss</AdminPassword>\r\n    <CodexAuthPath>C:\\ProgramData\\OutlookAI\\auth.json</CodexAuthPath>\r\n"
                    + "    <VoiceModel>gpt-realtime-2</VoiceModel>\r\n    <MaxBulkExportRows>500</MaxBulkExportRows>\r\n</Config>\r\n");
                File.WriteAllBytes(sandbox.ConfigPath, original);

                var r = sandbox.Run("1", "2", "5");

                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
                Assert.Equal(original, File.ReadAllBytes(Path.Combine(sandbox.BackupRoot, "config.xml.v1.backup.20260922-120000")));
                Assert.Equal("config.xml", r["left"]);
            }
        }

        [Fact]
        public void Update_StoppedAfterStep2_LeavesConfigInPlace_ForTheRerun()
        {
            using (var sandbox = new InstallSandbox())
            {
                var original = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><CodexAuthPath>C:\\x\\auth.json</CodexAuthPath><MaxBulkExportRows>500</MaxBulkExportRows></Config>");
                File.WriteAllBytes(sandbox.ConfigPath, original);

                sandbox.Run("1", "2");          // e.g. step 3 then fails to copy files
                sandbox.Run("1", "2", "5");     // the admin runs the installer again

                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
            }
        }

        [Fact]
        public void Update_AddsAMissingAuthPath_WrittenSoTheAddInReadsTheSamePassword()
        {
            using (var sandbox = new InstallSandbox())
            {
                File.WriteAllBytes(sandbox.ConfigPath, Windows1252.GetBytes(
                    "<?xml version=\"1.0\" encoding=\"windows-1252\"?><Config><AdminPassword>p\u00E4ss</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows></Config>"));

                var r = sandbox.Run("1", "2", "5");

                var root = XDocument.Load(sandbox.ConfigPath).Root;   // how Config.cs reads it
                Assert.Equal("p\u00E4ss", root.Element("AdminPassword").Value);
                Assert.Equal("750", root.Element("MaxBulkExportRows").Value);
                Assert.Equal(AuthPath, root.Element("CodexAuthPath").Value);
                Assert.Equal("config.xml", r["left"]);   // the staged config.xml.new was swapped in
            }
        }

        [Fact]
        public void Update_AFailedRewrite_LeavesTheOldConfigIntact()
        {
            using (var sandbox = new InstallSandbox())
            {
                var original = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
                File.WriteAllBytes(sandbox.ConfigPath, original);

                // A folder where step 5 stages the new file makes that write fail.
                var r = sandbox.Run("1", "2",
                    "New-Item -ItemType Directory -Path ($ConfigFilePath + '.new') | Out-Null",
                    "try { . (Get-InstallerStep 5) | Out-Null; 'step5=ok' } catch { 'step5=failed' }");

                Assert.Equal("failed", r["step5"]);
                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
            }
        }

        [Fact]
        public void Update_AFailedSwap_LeavesTheConfigInPlace_AndTheInstallCarriesOn()
        {
            using (var sandbox = new InstallSandbox())
            {
                var original = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
                File.WriteAllBytes(sandbox.ConfigPath, original);

                // A folder where the replaced file is set aside makes File.Replace fail.
                var r = sandbox.Run("1", "2",
                    "New-Item -ItemType Directory -Path ($ConfigFilePath + '.old\\x') -Force | Out-Null",
                    "try { . (Get-InstallerStep 5) | Out-Null; 'step5=ok' } catch { 'step5=failed' }");

                Assert.Equal("ok", r["step5"]);
                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
                Assert.False(File.Exists(sandbox.ConfigPath + ".new"));
            }
        }

        [Fact]
        public void Update_AReadOnlyConfig_IsLeftInPlace_AndTheInstallCarriesOn()
        {
            using (var sandbox = new InstallSandbox())
            {
                var original = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
                File.WriteAllBytes(sandbox.ConfigPath, original);
                File.SetAttributes(sandbox.ConfigPath, FileAttributes.ReadOnly);

                var r = sandbox.Run("1", "2", "try { . (Get-InstallerStep 5) | Out-Null; 'step5=ok' } catch { 'step5=failed' }");

                Assert.Equal("ok", r["step5"]);
                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
                Assert.Equal("config.xml", r["left"]);
            }
        }

        [Theory]
        [InlineData(true)]    // the new file was complete: the swap is finished
        [InlineData(false)]   // the new file was cut short: the replaced one comes back
        public void Rerun_AfterASwapStoppedHalfway_RestoresTheConfig(bool newFileComplete)
        {
            using (var sandbox = new InstallSandbox())
            {
                var newFile = Encoding.UTF8.GetBytes(newFileComplete
                    ? "<Config><AdminPassword>s3cret</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows><CodexAuthPath>C:\\x\\auth.json</CodexAuthPath></Config>"
                    : "<Config><AdminPassword>s3c");
                var oldFile = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><MaxBulkExportRows>750</MaxBulkExportRows></Config>");
                File.WriteAllBytes(sandbox.ConfigPath + ".new", newFile);
                File.WriteAllBytes(sandbox.ConfigPath + ".old", oldFile);

                var r = sandbox.Run("1", "2", "5");

                var root = XDocument.Load(sandbox.ConfigPath).Root;
                Assert.Equal("s3cret", root.Element("AdminPassword").Value);
                Assert.Equal("750", root.Element("MaxBulkExportRows").Value);
                if (newFileComplete) Assert.Equal(newFile, File.ReadAllBytes(sandbox.ConfigPath));
                Assert.Equal("config.xml", r["left"]);
            }
        }

        [Fact]
        public void Update_AConfigThatIsNotValidXml_IsKeptByteForByte()
        {
            using (var sandbox = new InstallSandbox())
            {
                var original = Encoding.UTF8.GetBytes("<Config><AdminPassword>s3cret</AdminPassword><Model>gpt-5.5</Config>");
                File.WriteAllBytes(sandbox.ConfigPath, original);

                sandbox.Run("1", "2", "5");

                Assert.Equal(original, File.ReadAllBytes(sandbox.ConfigPath));
            }
        }

        [Fact]
        public void FreshInstall_CarriesOverFromTheNewestBackup()
        {
            using (var sandbox = new InstallSandbox())
            {
                // Named by an installer before v2.2.1 run under the Thai calendar (2569 = 2026).
                sandbox.WriteBackup("25690922-090000", new DateTime(2026, 9, 22, 9, 0, 0, DateTimeKind.Utc),
                    "<Config><AdminPassword>older</AdminPassword><Model>gpt-5.4</Model></Config>");
                sandbox.WriteBackup("20260922-120000", new DateTime(2026, 9, 22, 12, 0, 0, DateTimeKind.Utc),
                    "<Config><AdminPassword>newest</AdminPassword><CodexAuthPath>C:\\x\\auth.json</CodexAuthPath><Model>gpt-6-sol</Model><VoiceModel>gpt-realtime-2</VoiceModel></Config>");
                sandbox.WriteBackup("20250101-000000", new DateTime(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc),
                    "<Config><AdminPassword>oldest</AdminPassword></Config>");

                var r = sandbox.Run("1", "2", "5");

                Assert.Equal(new[] { "AdminPassword", "CodexAuthPath", "Model" }, ElementNames(sandbox.ConfigPath));
                var root = XDocument.Load(sandbox.ConfigPath).Root;
                Assert.Equal("newest", root.Element("AdminPassword").Value);
                Assert.Equal(AuthPath, root.Element("CodexAuthPath").Value);
                Assert.Equal("gpt-6-sol", root.Element("Model").Value);
                Assert.Equal("config.xml", r["left"]);
            }
        }

        [Theory]
        [InlineData("<Config><AdminPassword>ab&#x1;cd</AdminPassword></Config>")]   // loads, but can't be written back
        [InlineData("<Config><AdminPass")]
        public void FreshInstall_ABackupItCannotUse_StillGetsTheDefaults(string newestBackup)
        {
            using (var sandbox = new InstallSandbox())
            {
                sandbox.WriteBackup("20991231-235959", DateTime.UtcNow, newestBackup);

                var r = sandbox.Run("1", "2", "try { . (Get-InstallerStep 5) | Out-Null; 'step5=ok' } catch { 'step5=failed' }");

                Assert.Equal("ok", r["step5"]);
                Assert.Equal(new[] { "AdminPassword", "CodexAuthPath" }, ElementNames(sandbox.ConfigPath));
                Assert.Equal("admin", XDocument.Load(sandbox.ConfigPath).Root.Element("AdminPassword").Value);
            }
        }

        [Fact]
        public void FreshInstall_IgnoresACutShortNewFile()
        {
            using (var sandbox = new InstallSandbox())
            {
                File.WriteAllText(sandbox.ConfigPath + ".new", "<Config><AdminPass");

                var r = sandbox.Run("1", "2", "5");

                Assert.Equal(new[] { "AdminPassword", "CodexAuthPath" }, ElementNames(sandbox.ConfigPath));
                Assert.Equal("admin", XDocument.Load(sandbox.ConfigPath).Root.Element("AdminPassword").Value);
                Assert.Equal("config.xml", r["left"]);
            }
        }

        [Fact]
        public void PerUserStep_RenamesOnlyV1Files()
        {
            using (var sandbox = new InstallSandbox())
            {
                var files = new Dictionary<string, string>
                {
                    ["alice"] = "<Config><ApiKey>x</ApiKey><OpenAIApiKey>y</OpenAIApiKey><AdminPassword>a</AdminPassword><Model>claude-opus-4-6</Model><MaxTokens>2048</MaxTokens></Config>",
                    ["bob"] = "<Config><AdminPassword>a</AdminPassword><Model>gpt-5.5</Model><ReasoningEffort>None</ReasoningEffort><WriteToolsEnabled>true</WriteToolsEnabled><EnabledWriteTools /></Config>",
                    ["carol"] = "<Config><AdminPass",
                    ["Public"] = "<Config><ApiKey>x</ApiKey></Config>",
                };
                foreach (var user in files.Keys)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(sandbox.UserConfigPath(user)));
                    File.WriteAllText(sandbox.UserConfigPath(user), files[user]);
                }

                sandbox.Run("7");

                Assert.False(File.Exists(sandbox.UserConfigPath("alice")));
                Assert.Equal(files["alice"], File.ReadAllText(sandbox.UserConfigPath("alice") + ".v1.backup.20260922-120000"));
                foreach (var kept in new[] { "bob", "carol", "Public" })
                {
                    Assert.Equal(files[kept], File.ReadAllText(sandbox.UserConfigPath(kept)));
                }
            }
        }

        [Fact]
        public void SharedFolderStep_GrantsAuthenticatedUsersModify_ByTheirSid()
        {
            // By SID, so it also works where the group's name is localized.
            using (var sandbox = new InstallSandbox())
            {
                var r = sandbox.Run("6",
                    "$rule = (Get-Acl -LiteralPath $ProgramDataPath).Access | Where-Object { -not $_.IsInherited -and $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]).Value -eq 'S-1-5-11' } | Select-Object -First 1",
                    "'rights=' + $rule.FileSystemRights",
                    "'inherits=' + $rule.InheritanceFlags");

                Assert.Contains("Modify", r["rights"]);
                Assert.Contains("ContainerInherit", r["inherits"]);
                Assert.Contains("ObjectInherit", r["inherits"]);
            }
            var src = File.ReadAllText(RepoFiles.Find("Deploy", "Install-OutlookAI.ps1"));
            var step6 = src.Substring(src.IndexOf("# --- 6.", StringComparison.Ordinal));
            Assert.Contains("/grant \"*S-1-5-11:(OI)(CI)M\"", step6.Substring(0, step6.IndexOf("# --- 7.", StringComparison.Ordinal)));
        }

        [Fact]
        public void Timestamp_IsGregorian_WhateverCalendarTheAdminUses()
        {
            var r = RunInstallerHelpers(string.Join("\n",
                "$line = $installerText -split \"`r?`n\" | Where-Object { $_ -match '^\\$Timestamp\\s*=' } | Select-Object -First 1",
                "& {",
                "    [System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::GetCultureInfo('th-TH')",
                "    'culture-year=' + (Get-Date -Format 'yyyy')",
                "    Invoke-Expression $line",
                "    'timestamp=' + $Timestamp",
                "}"));

            Assert.StartsWith("25", r["culture-year"]);   // Thai solar calendar in effect
            Assert.StartsWith(DateTime.Now.Year.ToString(System.Globalization.CultureInfo.InvariantCulture), r["timestamp"]);
        }
    }
}
