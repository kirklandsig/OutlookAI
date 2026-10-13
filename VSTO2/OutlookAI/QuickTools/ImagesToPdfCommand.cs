using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Diagnostics;
using OutlookAI.Services;
using OutlookAI.Services.Export;
using OutlookAI.Services.QuickTools;
using OutlookAI.Services.Updates;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI.QuickTools
{
    /// <summary>Quick Tools -> Clean up images -> PDF, from the Inbox or an opened email.</summary>
    internal static class ImagesToPdfCommand
    {
        private const string Caption = "Clean up images → PDF";
        private const string NoImages = "No images found in this email.";
        private const string NoImagesRichText = "No images found. Images pasted into Rich Text emails aren't supported yet.";

        // The progress or review window of the run in progress (one run at a time).
        private static Form _active;

        public static async void Run(object context)
        {
            if (_active != null)
            {
                try { _active.Activate(); } catch (ObjectDisposedException) { }
                return;
            }
            if (SynchronizationContext.Current == null)
            {
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            }

            var owner = OwnerWindow.For(context);
            string runFolder = null;
            try
            {
                var mail = FindMail(context, owner);
                if (mail == null) return;

                var addIn = Globals.ThisAddIn;
                var auth = addIn?.AuthService;
                var chat = addIn?.ChatService;
                if (auth == null || chat == null)
                {
                    Info(owner, "OutlookAI isn't ready. Restart Outlook and try again.");
                    return;
                }
                var authStatus = auth.GetStatus();
                if (authStatus.State != AuthState.Authenticated)
                {
                    Info(owner, authStatus.State == AuthState.Error
                        ? "OutlookAI couldn't read its ChatGPT sign-in: " + authStatus.Message + "\n\n" + CodexAuthService.NotSignedInMessage
                        : CodexAuthService.NotSignedInMessage);
                    return;
                }

                var subject = Safe(() => mail.Subject) ?? "";
                var sender = Safe(() => mail.SenderName) ?? "";
                runFolder = QuickToolsRunFolder.Create(DateTime.UtcNow);
                var collected = EmailImageCollector.Collect(mail, runFolder);
                if (collected.Images.Count == 0)
                {
                    Info(owner, collected.IsRichText ? NoImagesRichText : NoImages);
                    return;
                }

                PipelineResult result;
                using (var cts = new CancellationTokenSource())
                using (var progressForm = new QuickToolsProgressForm(cts))
                {
                    _active = progressForm;
                    progressForm.Show(owner);
                    var progress = new Progress<string>(progressForm.SetStatus);
                    var planner = new PagePlanner(chat);
                    var pipeline = new ImagesToPdfPipeline(planner.PlanAsync, runFolder);
                    var started = Stopwatch.StartNew();
                    try
                    {
                        result = await Task.Run(() => pipeline.RunAsync(collected.Images, progress, cts.Token));
                        if (cts.IsCancellationRequested)
                        {
                            TraceLog.Write("QuickTools: cancelled", "QuickTools");
                            return;
                        }
                    }
                    catch (OperationCanceledException) when (cts.IsCancellationRequested)
                    {
                        TraceLog.Write("QuickTools: cancelled", "QuickTools");
                        return;
                    }
                    finally
                    {
                        _active = null;
                        progressForm.Finish();
                    }
                    TraceLog.Write("QuickTools: " + result.Pages.Count + " pages, " + result.SkippedSmall + " small skipped, "
                        + started.Elapsed.TotalSeconds.ToString("F1", CultureInfo.InvariantCulture) + " s", "QuickTools");
                }

                if (result.Pages.Count == 0)
                {
                    Info(owner, NoImages + (result.SkippedSmall > 0 ? " (" + result.SkippedSmall + " small images skipped.)" : ""));
                    return;
                }

                var note = result.NoneAnalyzed
                    ? "ChatGPT couldn't analyze these images, so they have basic cleanup only. Check each page before saving."
                    : null;
                var heading = "\"" + subject + "\" from " + sender;
                var folder = runFolder;
                using (var review = new ImagesToPdfReviewForm(heading, note, result.Pages,
                    page => Task.Run(() => ImagesToPdfPipeline.Reclean(page)),
                    (window, pages) => SavePdf(window, pages, subject, sender, folder)))
                {
                    _active = review;
                    try
                    {
                        review.ShowDialog(owner);
                    }
                    finally
                    {
                        _active = null;
                    }
                }
            }
            catch (NotSignedInException ex)
            {
                Info(owner, ex.Message);
            }
            catch (Exception ex)
            {
                TraceLog.Write("QuickTools: failed (" + ex.GetType().Name + ")", "QuickTools");
                MessageBox.Show(owner, "Error: " + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _active = null;
                if (runFolder != null) QuickToolsRunFolder.Delete(runFolder);
            }
        }

        private static Outlook.MailItem FindMail(object context, IWin32Window owner)
        {
            object item = null;
            var explorer = context as Outlook.Explorer;
            if (explorer != null)
            {
                Outlook.Selection selection = null;
                try { selection = explorer.Selection; } catch (COMException) { }
                if (selection == null || selection.Count != 1)
                {
                    Info(owner, "Select one email first.");
                    return null;
                }
                item = selection[1];
            }
            else
            {
                var inspector = context as Outlook.Inspector;
                if (inspector != null)
                {
                    try { item = inspector.CurrentItem; } catch (COMException) { }
                }
            }

            var mail = item as Outlook.MailItem;
            if (mail == null) Info(owner, item == null ? "Select one email first." : "This works on email messages.");
            return mail;
        }

        private static bool SavePdf(IWin32Window owner, System.Collections.Generic.IReadOnlyList<ReviewPage> pages, string subject, string sender, string runFolder)
        {
            var resolver = Globals.ThisAddIn?.ExportPathResolver ?? new ExportPathResolver();
            string folder = null;
            try
            {
                resolver.EnsureExists();
                folder = resolver.ResolveBaseDir();
            }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }

            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "Save PDF";
                dialog.Filter = "PDF files (*.pdf)|*.pdf";
                dialog.DefaultExt = "pdf";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = true;
                if (folder != null) dialog.InitialDirectory = folder;
                dialog.FileName = ExportFilenameSanitizer.Build(sender + " - " + subject, ".pdf", DateTimeOffset.Now,
                    name => folder != null && File.Exists(Path.Combine(folder, name)));
                if (dialog.ShowDialog(owner) != DialogResult.OK) return false;

                try
                {
                    var temp = Path.Combine(runFolder, "output.pdf");
                    ImagePdfWriter.WriteFile(ImagesToPdfPipeline.FinalJpegs(pages, runFolder), subject, Producer(), temp);
                    var error = ImagesToPdfOutput.Publish(temp, dialog.FileName);
                    if (error != null)
                    {
                        MessageBox.Show(owner, error, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    TraceLog.Write("QuickTools: save failed (" + ex.GetType().Name + ")", "QuickTools");
                    MessageBox.Show(owner, "Error: " + ex.Message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                TraceLog.Write("QuickTools: saved " + pages.Count + " pages", "QuickTools");
                try
                {
                    Process.Start(new ProcessStartInfo(dialog.FileName) { UseShellExecute = true });
                }
                catch (Exception ex) when (ex is Win32Exception || ex is InvalidOperationException)
                {
                    Info(owner, "Saved to " + dialog.FileName + ", but it couldn't be opened automatically.");
                }
                return true;
            }
        }

        private static string Producer()
        {
            try
            {
                var tag = UpdateManifest.LoadFromInstallDir().Tag;
                return string.IsNullOrEmpty(tag) ? "OutlookAI" : "OutlookAI " + tag;
            }
            catch (Exception)
            {
                return "OutlookAI";
            }
        }

        private static void Info(IWin32Window owner, string message) =>
            MessageBox.Show(owner, message, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

        private static T Safe<T>(Func<T> read)
        {
            try
            {
                return read();
            }
            catch (COMException)
            {
                return default(T);
            }
        }
    }
}
