using System;
using System.Collections.Generic;
using System.IO;
using OutlookAI.Services.Updates;
using Xunit;

namespace OutlookAI.Tests.Services.Updates
{
    public class UpdateHistoryLogTests : IDisposable
    {
        private readonly string _path;

        public UpdateHistoryLogTests()
        {
            var dir = Path.Combine(Path.GetTempPath(), "updater-history-tests", Path.GetRandomFileName());
            Directory.CreateDirectory(dir);
            _path = Path.Combine(dir, "update-history.json");
        }

        public void Dispose()
        {
            try { var d = Path.GetDirectoryName(_path); if (d != null) Directory.Delete(d, true); } catch { }
        }

        [Fact]
        public void Append_WritesEntryAndReadAllReturnsIt()
        {
            var log = new UpdateHistoryLog(_path);
            log.Append("check", "newer_available", "v2.1.0", "");

            var entries = log.ReadAll();
            Assert.Single(entries);
            Assert.Equal("check", entries[0].Action);
            Assert.Equal("newer_available", entries[0].Result);
            Assert.Equal("v2.1.0", entries[0].Tag);
        }

        [Fact]
        public void Append_KeepsOnlyLast50Entries()
        {
            var log = new UpdateHistoryLog(_path);
            for (var i = 0; i < 60; i++) log.Append("check", "noop", "v" + i, "");

            var entries = log.ReadAll();
            Assert.Equal(50, entries.Count);
            Assert.Equal("v10", entries[0].Tag);  // oldest 10 dropped
            Assert.Equal("v59", entries[49].Tag);
        }

        [Fact]
        public void ReadAll_MissingFile_ReturnsEmpty()
        {
            var log = new UpdateHistoryLog(_path);
            Assert.Empty(log.ReadAll());
        }

        [Fact]
        public void ReadAll_MalformedFile_ReturnsEmptyAndDoesNotThrow()
        {
            File.WriteAllText(_path, "not json");
            var log = new UpdateHistoryLog(_path);
            Assert.Empty(log.ReadAll());
        }
    }
}
