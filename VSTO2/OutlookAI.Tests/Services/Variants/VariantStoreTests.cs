using System;
using OutlookAI.Services.Variants;
using Xunit;

namespace OutlookAI.Tests.Services.Variants
{
    public class VariantStoreTests
    {
        private static Variant V(Tone t, string body) =>
            new Variant { Tone = t, Body = body };

        [Fact]
        public void Replace_ReplacesAllExistingVariants()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "a"), V(Tone.Brief, "b") });
            Assert.Equal(2, s.Count);

            s.Replace(new[] { V(Tone.Direct, "c") });

            Assert.Equal(1, s.Count);
            var snap = s.Snapshot();
            Assert.Equal(Tone.Direct, snap[0].Tone);
            Assert.Equal("c", snap[0].Body);
        }

        [Fact]
        public void Update_ReplacesSingleIndexInPlace()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "a"), V(Tone.Brief, "b"), V(Tone.Direct, "c") });

            s.Update(1, V(Tone.Persuasive, "B-prime"));

            var snap = s.Snapshot();
            Assert.Equal(3, snap.Count);
            Assert.Equal(Tone.Formal, snap[0].Tone);
            Assert.Equal(Tone.Persuasive, snap[1].Tone);
            Assert.Equal("B-prime", snap[1].Body);
            Assert.Equal(Tone.Direct, snap[2].Tone);
        }

        [Fact]
        public void Update_ThrowsOnOutOfRangeIndex()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "a") });
            Assert.Throws<ArgumentOutOfRangeException>(() => s.Update(-1, V(Tone.Brief, "x")));
            Assert.Throws<ArgumentOutOfRangeException>(() => s.Update(5, V(Tone.Brief, "x")));
        }

        [Fact]
        public void Clear_EmptiesTheStore()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "a"), V(Tone.Brief, "b") });

            s.Clear();

            Assert.Equal(0, s.Count);
            Assert.Empty(s.Snapshot());
        }

        [Fact]
        public void Snapshot_ReturnsIndependentCopy()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "a") });

            var snap = s.Snapshot();
            s.Clear();

            // Snapshot still has the original entry; not affected by Clear.
            Assert.Single(snap);
            Assert.Equal("a", snap[0].Body);
        }

        [Fact]
        public void TwoStores_AreIsolated()
        {
            var a = new VariantStore();
            var b = new VariantStore();
            a.Replace(new[] { V(Tone.Formal, "a-val") });
            b.Replace(new[] { V(Tone.Brief, "b-val"), V(Tone.Direct, "b-val-2") });

            Assert.Single(a.Snapshot());
            Assert.Equal(2, b.Count);
            Assert.Equal("a-val", a.Snapshot()[0].Body);
            Assert.Equal("b-val", b.Snapshot()[0].Body);
        }

        [Fact]
        public void Replace_NullArgument_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => new VariantStore().Replace(null));
        }

        [Fact]
        public void Update_NullVariant_Throws()
        {
            var s = new VariantStore();
            s.Replace(new[] { V(Tone.Formal, "x") });
            Assert.Throws<ArgumentNullException>(() => s.Update(0, null));
        }
    }
}
