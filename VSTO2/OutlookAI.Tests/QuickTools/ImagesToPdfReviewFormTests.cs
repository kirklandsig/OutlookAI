using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.QuickTools;
using OutlookAI.Services.QuickTools;
using OutlookAI.Tests.Helpers;
using Xunit;

namespace OutlookAI.Tests.QuickTools
{
    public class ImagesToPdfReviewFormTests
    {
        private static ReviewPage Page(int number, string kind = "document_page", bool readable = true)
        {
            var page = new ReviewPage(number, "image" + number + ".jpg")
            {
                Readable = readable,
                Plan = PagePlan.FromJson("{\"kind\":\"" + kind + "\",\"reason\":\"company logo\",\"corners\":[[0,0],[1,0],[1,1],[0,1]],\"color\":\"grayscale\",\"strength\":\"normal\"}"),
                UprightPath = "upright.jpg",
                CleanedPath = "cleaned.jpg",
            };
            page.Include = readable && page.Plan.Kind == PageKind.DocumentPage;
            return page;
        }

        private static ImagesToPdfReviewForm NewForm(IReadOnlyList<ReviewPage> pages,
            System.Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> save = null,
            System.Func<ReviewPage, Task> reclean = null) =>
            new ImagesToPdfReviewForm("\"Lease docs\" from Jane Example", null, pages,
                reclean ?? (p => Task.CompletedTask), save ?? ((o, p) => true), (path, size) => new Bitmap(20, 30));

        [Fact]
        public void Opens_WithDocumentsIncluded_AndOthersNot()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2, "not_document") }))
                {
                    Assert.True(form.Pages[0].Include);
                    Assert.False(form.Pages[1].Include);
                    Assert.True(form.CanSave);
                }
            });
        }

        [Fact]
        public void Save_IsDisabled_WhenNoPageIsIncluded()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2) }))
                {
                    form.SetIncluded(0, false);
                    form.SetIncluded(1, false);

                    Assert.False(form.CanSave);
                }
            });
        }

        [Fact]
        public void UnreadablePage_CantBeIncluded()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1, readable: false) }))
                {
                    form.SetIncluded(0, true);

                    Assert.False(form.Pages[0].Include);
                    Assert.False(form.CanSave);
                }
            });
        }

        [Fact]
        public void RotateAndOriginal_ApplyToTheSelectedPageOnly()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1), Page(2) }))
                {
                    form.SelectPage(0);
                    form.Rotate(1);
                    form.SetUseOriginal(true);

                    Assert.Equal(1, form.Pages[0].QuarterTurns);
                    Assert.True(form.Pages[0].UseOriginal);
                    Assert.Equal(0, form.Pages[1].QuarterTurns);
                    Assert.False(form.Pages[1].UseOriginal);
                    form.Rotate(3);
                    Assert.Equal(0, form.Pages[0].QuarterTurns);
                }
            });
        }

        // Its cleaned file was never written: Cleaned and a color change would point at nothing.
        [Fact]
        public void CleanupFailedPage_StaysOnItsOriginal()
        {
            Sta.Run(() =>
            {
                var recleaned = 0;
                var page = Page(1);
                page.CleanupFailed = true;
                page.UseOriginal = true;
                using (var form = NewForm(new[] { page }, reclean: p => { recleaned++; return Task.CompletedTask; }))
                {
                    form.SelectPage(0);
                    form.SetUseOriginal(false);
                    form.SetColorAsync(1).GetAwaiter().GetResult();

                    Assert.True(form.Pages[0].UseOriginal);
                    Assert.Equal(0, recleaned);
                }
            });
        }

        [Fact]
        public void ColorChange_RecleansThatPage()
        {
            Sta.Run(() =>
            {
                var recleaned = new List<ReviewPage>();
                using (var form = NewForm(new[] { Page(1), Page(2) }, reclean: p => { recleaned.Add(p); return Task.CompletedTask; }))
                {
                    form.SelectPage(1);
                    form.SetColorAsync(1).GetAwaiter().GetResult();
                    form.SetColorAsync(1).GetAwaiter().GetResult(); // no change: no second clean

                    Assert.Equal(PageColor.Color, form.Pages[1].ColorOverride);
                    Assert.Same(form.Pages[1], Assert.Single(recleaned));
                }
            });
        }

        [Fact]
        public void ColorChange_FailedReclean_RevertsTheChoice()
        {
            Sta.Run(() =>
            {
                var original = ImagesToPdfReviewForm.ShowWarning;
                var warned = 0;
                ImagesToPdfReviewForm.ShowWarning = (owner, text) => warned++;
                try
                {
                    using (var form = NewForm(new[] { Page(1) },
                        reclean: p => Task.FromException(new System.IO.IOException("x"))))
                    {
                        form.SelectPage(0);
                        form.ChooseColor(1);

                        Assert.Null(form.Pages[0].ColorOverride);
                        Assert.Equal(0, form.ColorIndex);
                        Assert.Equal(1, warned);
                    }
                }
                finally
                {
                    ImagesToPdfReviewForm.ShowWarning = original;
                }
            });
        }

        [Fact]
        public void Save_PassesTheIncludedPagesInOrder_AndCloses()
        {
            Sta.Run(() =>
            {
                IReadOnlyList<ReviewPage> saved = null;
                var a = Page(1);
                var b = Page(2);
                var c = Page(3, "not_document");
                using (var form = NewForm(new[] { a, b, c }, save: (owner, pages) => { saved = pages; return true; }))
                {
                    form.SelectPage(1);
                    form.MovePage(-1);
                    form.SavePdf();

                    Assert.Equal(new[] { b, a }, saved.ToArray());
                    Assert.Equal(DialogResult.OK, form.DialogResult);
                }
            });
        }

        [Fact]
        public void Save_StaysOpen_WhenSavingFails()
        {
            Sta.Run(() =>
            {
                using (var form = NewForm(new[] { Page(1) }, save: (owner, pages) => false))
                {
                    form.SavePdf();

                    Assert.Equal(DialogResult.None, form.DialogResult);
                    Assert.False(form.IsDisposed);
                }
            });
        }
    }
}
