using System;
using System.IO;
using System.Threading;

namespace OutlookAI.Diagnostics
{
    /// <summary>
    /// File-based trace logger for diagnosing freezes / hangs in the
    /// add-in. Writes timestamped entries with thread context to
    /// <c>%LOCALAPPDATA%\OutlookAI\trace.log</c>. Append-only, single
    /// process, fail-silently so a logging hiccup never propagates.
    ///
    /// Capture flow during a freeze investigation:
    ///   1. Reproduce the freeze.
    ///   2. Read %LOCALAPPDATA%\OutlookAI\trace.log.
    ///   3. The last line written tells you where the UI thread
    ///      stopped doing work.
    ///
    /// Disabled when <c>Config.TraceEnabled</c> is false (default true
    /// for Phase 2 dogfood; flip to false for production once stable).
    /// </summary>
    public static class TraceLog
    {
        private static readonly object _lock = new object();
        private static readonly string _logPath;
        private static int _uiThreadId;
        private static readonly DateTime _processStart = DateTime.UtcNow;

        static TraceLog()
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookAI");
                Directory.CreateDirectory(dir);
                _logPath = Path.Combine(dir, "trace.log");
                // Truncate on each process start so logs reflect this session.
                File.WriteAllText(_logPath,
                    "=== OutlookAI trace start "
                    + DateTime.UtcNow.ToString("o")
                    + " PID=" + System.Diagnostics.Process.GetCurrentProcess().Id
                    + " ===" + Environment.NewLine);
            }
            catch
            {
                _logPath = null;
            }
        }

        /// <summary>
        /// Mark the current thread as the UI thread. Call once from
        /// ThisAddIn_Startup. Subsequent <see cref="Write"/> calls tag
        /// log lines with whether they fired on the UI thread or not.
        /// </summary>
        public static void MarkUiThread()
        {
            _uiThreadId = Thread.CurrentThread.ManagedThreadId;
            Write("UI thread marker installed", "TraceLog");
        }

        public static void Write(string message, string source = "")
        {
            if (_logPath == null) return;
            try
            {
                var tid = Thread.CurrentThread.ManagedThreadId;
                var ui = (_uiThreadId != 0 && tid == _uiThreadId) ? "UI" : "  ";
                var elapsed = (DateTime.UtcNow - _processStart).TotalMilliseconds;
                var line = string.Format(
                    "[{0,9:F1}ms T{1,4} {2}] {3}{4}{5}",
                    elapsed,
                    tid,
                    ui,
                    string.IsNullOrEmpty(source) ? "" : "(" + source + ") ",
                    message,
                    Environment.NewLine);
                lock (_lock)
                {
                    File.AppendAllText(_logPath, line);
                }
            }
            catch
            {
                // Swallow: tracing should never break the add-in.
            }
        }

        /// <summary>
        /// Write an "entering" line and return an <see cref="IDisposable"/>
        /// that writes the matching "exiting" line plus elapsed ms when
        /// disposed. Use with <c>using (TraceLog.Scope("...")){...}</c>.
        /// </summary>
        public static IDisposable Scope(string label, string source = "")
        {
            Write(">> " + label, source);
            var sw = System.Diagnostics.Stopwatch.StartNew();
            return new ScopeMarker(label, source, sw);
        }

        private sealed class ScopeMarker : IDisposable
        {
            private readonly string _label;
            private readonly string _source;
            private readonly System.Diagnostics.Stopwatch _sw;
            private bool _disposed;
            public ScopeMarker(string label, string source, System.Diagnostics.Stopwatch sw)
            {
                _label = label;
                _source = source;
                _sw = sw;
            }
            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _sw.Stop();
                Write("<< " + _label + " (" + _sw.ElapsedMilliseconds + "ms)", _source);
            }
        }
    }
}
