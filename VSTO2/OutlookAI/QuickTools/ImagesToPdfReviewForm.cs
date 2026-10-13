using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services.QuickTools;

namespace OutlookAI.QuickTools
{
    /// <summary>
    /// The pages before they become a PDF: include or drop, reorder, rotate,
    /// use the original, change color. Holds file paths, 160 px thumbnails and
    /// one preview, never full-size images.
    /// </summary>
    internal sealed class ImagesToPdfReviewForm : Form
    {
        internal const int ThumbSize = 160;
        internal const int PreviewSize = 1200;
        private const string Caption = "Clean up images \u2192 PDF";

        internal static Action<IWin32Window, string> ShowWarning = (owner, text) =>
            MessageBox.Show(owner, text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private readonly List<ReviewPage> _pages;
        private readonly Func<ReviewPage, Task> _reclean;
        private readonly Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> _save;
        private readonly Func<string, int, Bitmap> _loadScaled;

        private readonly ListView _list;
        private readonly ImageList _thumbs;
        private readonly PictureBox _preview;
        private readonly Button _up;
        private readonly Button _down;
        private readonly Button _rotateLeft;
        private readonly Button _rotateRight;
        private readonly Button _cancel;
        private readonly Button _savePdf;
        private readonly RadioButton _cleaned;
        private readonly RadioButton _original;
        private readonly ComboBox _color;
        private int _selected = -1;
        private bool _updating;
        private bool _busy;

        public ImagesToPdfReviewForm(
            string heading,
            string note,
            IReadOnlyList<ReviewPage> pages,
            Func<ReviewPage, Task> reclean,
            Func<IWin32Window, IReadOnlyList<ReviewPage>, bool> save,
            Func<string, int, Bitmap> loadScaled = null)
        {
            _pages = pages.ToList();
            _reclean = reclean;
            _save = save;
            _loadScaled = loadScaled ?? JpegFiles.LoadScaled;

            Text = Caption + " \u2014 " + heading;
            Size = new Size(960, 680);
            MinimumSize = new Size(760, 520);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.FromArgb(250, 249, 248);
            ShowInTaskbar = false;
            MinimizeBox = false;

            _thumbs = new ImageList { ImageSize = new Size(ThumbSize, ThumbSize), ColorDepth = ColorDepth.Depth24Bit };
            // Creating the handle now makes the list copy every bitmap it is given; before
            // that it keeps the caller's reference, which the callers here dispose.
            GC.KeepAlive(_thumbs.Handle);
            _list = new ListView
            {
                Dock = DockStyle.Left,
                Width = 250,
                View = View.LargeIcon,
                CheckBoxes = true,
                MultiSelect = false,
                HideSelection = false,
                LargeImageList = _thumbs,
            };
            _preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(64, 64, 64) };

            var pageControls = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6), WrapContents = false };
            _up = MakeButton("Move up", 80);
            _down = MakeButton("Move down", 90);
            _rotateLeft = MakeButton("Rotate left", 90);
            _rotateRight = MakeButton("Rotate right", 95);
            _cleaned = new RadioButton { Text = "Cleaned", AutoSize = true, Checked = true, Margin = new Padding(16, 8, 0, 0) };
            _original = new RadioButton { Text = "Original", AutoSize = true, Margin = new Padding(8, 8, 0, 0) };
            var colorLabel = new Label { Text = "Color:", AutoSize = true, Margin = new Padding(16, 10, 0, 0) };
            _color = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100, Margin = new Padding(4, 6, 0, 0) };
            _color.Items.AddRange(new object[] { "Auto", "Color", "Grayscale" });
            pageControls.Controls.AddRange(new Control[] { _up, _down, _rotateLeft, _rotateRight, _cleaned, _original, colorLabel, _color });

            var dialogButtons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(6), FlowDirection = FlowDirection.RightToLeft };
            _savePdf = MakeButton("Save PDF\u2026", 110);
            _cancel = MakeButton("Cancel", 90);
            dialogButtons.Controls.AddRange(new Control[] { _savePdf, _cancel });

            Controls.Add(_preview);
            Controls.Add(_list);
            Controls.Add(pageControls);
            Controls.Add(dialogButtons);
            if (!string.IsNullOrEmpty(note))
            {
                Controls.Add(new Label { Text = note, Dock = DockStyle.Top, Height = 28, ForeColor = Color.DarkRed, Padding = new Padding(8, 6, 8, 0) });
            }

            foreach (var page in _pages)
            {
                using (var thumb = Thumbnail(page))
                {
                    _thumbs.Images.Add(Key(page), thumb);
                }
            }
            AddItems();

            _list.ItemCheck += (s, e) =>
            {
                if (_updating) return;
                if (!((ReviewPage)_list.Items[e.Index].Tag).Readable) e.NewValue = CheckState.Unchecked;
            };
            _list.ItemChecked += (s, e) =>
            {
                if (_updating) return;
                ((ReviewPage)e.Item.Tag).Include = e.Item.Checked;
                UpdateButtons();
            };
            _list.SelectedIndexChanged += (s, e) =>
            {
                if (_updating || _list.SelectedIndices.Count == 0) return;
                SelectPage(_list.SelectedIndices[0]);
            };
            _up.Click += (s, e) => MovePage(-1);
            _down.Click += (s, e) => MovePage(1);
            _rotateLeft.Click += (s, e) => Rotate(3);
            _rotateRight.Click += (s, e) => Rotate(1);
            _cleaned.CheckedChanged += (s, e) => { if (!_updating && _cleaned.Checked) SetUseOriginal(false); };
            _original.CheckedChanged += (s, e) => { if (!_updating && _original.Checked) SetUseOriginal(true); };
            _color.SelectedIndexChanged += async (s, e) => { if (!_updating) await SetColorAsync(_color.SelectedIndex); };
            _savePdf.Click += (s, e) => SavePdf();
            _cancel.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            FormClosing += (s, e) => { if (_busy) e.Cancel = true; };

            if (_pages.Count > 0) SelectPage(0);
            UpdateButtons();
        }

        internal IReadOnlyList<ReviewPage> Pages => _pages;

        internal IReadOnlyList<ListViewItem> ListItems => _list.Items.Cast<ListViewItem>().ToList();

        internal bool CanSave => _savePdf.Enabled;

        internal int ColorIndex => _color.SelectedIndex;

        internal void ChooseColor(int index) => _color.SelectedIndex = index;

        private ReviewPage Selected => _selected >= 0 && _selected < _pages.Count ? _pages[_selected] : null;

        internal void SelectPage(int index)
        {
            if (index < 0 || index >= _pages.Count) return;
            _selected = index;
            var page = _pages[index];
            _updating = true;
            try
            {
                _list.Items[index].Selected = true;
            }
            finally
            {
                _updating = false;
            }
            SyncControls(page);
            ShowPreview(page);
            UpdateButtons();
        }

        private void SyncControls(ReviewPage page)
        {
            _updating = true;
            try
            {
                _cleaned.Checked = !page.UseOriginal;
                _original.Checked = page.UseOriginal;
                _color.SelectedIndex = page.ColorOverride == null ? 0 : page.ColorOverride == PageColor.Color ? 1 : 2;
            }
            finally
            {
                _updating = false;
            }
        }

        internal void SetIncluded(int index, bool include)
        {
            var page = _pages[index];
            page.Include = include && page.Readable;
            _updating = true;
            try
            {
                _list.Items[index].Checked = page.Include;
            }
            finally
            {
                _updating = false;
            }
            UpdateButtons();
        }

        internal void MovePage(int delta)
        {
            var from = _selected;
            var to = from + delta;
            if (from < 0 || to < 0 || to >= _pages.Count) return;
            var page = _pages[from];
            _pages.RemoveAt(from);
            _pages.Insert(to, page);
            // The icon view keeps each item where it was drawn, so moving one item
            // changes the index order but not the screen; re-adding lays them out anew.
            _updating = true;
            _list.BeginUpdate();
            try
            {
                _list.Items.Clear();
                AddItems();
            }
            finally
            {
                _list.EndUpdate();
                _updating = false;
            }
            SelectPage(to);
            _list.Items[to].EnsureVisible();
        }

        // Thumbnails stay keyed by page number; the text shows the position.
        private void AddItems()
        {
            for (var i = 0; i < _pages.Count; i++)
            {
                var page = _pages[i];
                _list.Items.Add(new ListViewItem((i + 1) + ". " + page.Status, Key(page)) { Checked = page.Include, Tag = page });
            }
        }

        internal void Rotate(int quarterTurns)
        {
            var page = Selected;
            if (page == null || !page.Readable) return;
            page.QuarterTurns = (page.QuarterTurns + quarterTurns) % 4;
            RefreshPage(page);
        }

        internal void SetUseOriginal(bool useOriginal)
        {
            var page = Selected;
            if (page == null || !page.Readable || page.UseOriginal == useOriginal || (page.CleanupFailed && !useOriginal)) return;
            page.UseOriginal = useOriginal;
            RefreshPage(page);
        }

        internal async Task SetColorAsync(int choice)
        {
            var page = Selected;
            if (page == null || !page.Readable || page.CleanupFailed) return;
            var wanted = choice == 1 ? (PageColor?)PageColor.Color : choice == 2 ? PageColor.Grayscale : (PageColor?)null;
            if (wanted == page.ColorOverride) return;
            var previous = page.ColorOverride;
            var before = page.EffectiveColor;
            page.ColorOverride = wanted;
            if (page.EffectiveColor == before) return;

            _busy = true;
            UpdateButtons();
            try
            {
                await _reclean(page);
            }
            catch (Exception ex)
            {
                page.ColorOverride = previous;
                if (!IsDisposed)
                {
                    SyncControls(page);
                    ShowWarning(this, "This page couldn't be redone in that color.\n\n" + ex.Message);
                }
            }
            finally
            {
                _busy = false;
                if (!IsDisposed) UpdateButtons();
            }
            if (IsDisposed) return;
            RefreshPage(page);
        }

        internal void SavePdf()
        {
            var chosen = _pages.Where(p => p.Include && p.Readable).ToList();
            if (chosen.Count == 0 || _busy) return;
            if (_save(this, chosen))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _preview.Image?.Dispose();
                _thumbs.Dispose();
            }
            base.Dispose(disposing);
        }

        // Replaces the thumbnail in place: removing and re-adding would shift the
        // image indices the other list items resolved.
        private void RefreshPage(ReviewPage page)
        {
            var slot = _thumbs.Images.IndexOfKey(Key(page));
            if (slot >= 0)
            {
                using (var thumb = Thumbnail(page))
                {
                    _thumbs.Images[slot] = thumb;
                }
            }
            _list.Invalidate();
            if (_pages.IndexOf(page) == _selected) ShowPreview(page);
        }

        private void ShowPreview(ReviewPage page)
        {
            var old = _preview.Image;
            _preview.Image = Render(page, PreviewSize);
            old?.Dispose();
        }

        private Bitmap Thumbnail(ReviewPage page)
        {
            var thumb = new Bitmap(ThumbSize, ThumbSize);
            using (var g = Graphics.FromImage(thumb))
            {
                g.Clear(Color.Gainsboro);
                using (var image = Render(page, ThumbSize))
                {
                    if (image != null)
                    {
                        var scale = Math.Min((double)ThumbSize / image.Width, (double)ThumbSize / image.Height);
                        var w = (int)(image.Width * scale);
                        var h = (int)(image.Height * scale);
                        g.DrawImage(image, (ThumbSize - w) / 2, (ThumbSize - h) / 2, w, h);
                    }
                }
            }
            return thumb;
        }

        private Bitmap Render(ReviewPage page, int maxEdge)
        {
            if (!page.Readable) return null;
            var path = page.UseOriginal ? page.UprightPath : page.CleanedPath;
            try
            {
                var bitmap = _loadScaled(path, maxEdge);
                try
                {
                    bitmap.RotateFlip(PageCleaner.QuarterTurns(page.QuarterTurns));
                }
                catch
                {
                    bitmap.Dispose();
                    throw;
                }
                return bitmap;
            }
            catch (Exception ex) when (ex is System.IO.IOException || ex is ArgumentException || ex is OutOfMemoryException)
            {
                return null;
            }
        }

        private void UpdateButtons()
        {
            var page = Selected;
            var editable = !_busy && page != null && page.Readable;
            _cancel.Enabled = !_busy;
            _up.Enabled = !_busy && _selected > 0;
            _down.Enabled = !_busy && _selected >= 0 && _selected < _pages.Count - 1;
            _rotateLeft.Enabled = _rotateRight.Enabled = _original.Enabled = editable;
            _cleaned.Enabled = _color.Enabled = editable && !page.CleanupFailed;
            _savePdf.Enabled = !_busy && _pages.Any(p => p.Include && p.Readable);
        }

        private static string Key(ReviewPage page) => page.Number.ToString(System.Globalization.CultureInfo.InvariantCulture);

        private static Button MakeButton(string text, int width) => new Button
        {
            Text = text,
            Width = width,
            Height = 28,
            ForeColor = Color.Black,
            BackColor = SystemColors.ButtonFace,
            UseVisualStyleBackColor = false,
            Margin = new Padding(4),
        };
    }
}
