using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutlookAI.Services;
using OutlookAI.Services.Chat;
using OutlookAI.Services.Tools;
using OutlookAI.Services.Variants;

namespace OutlookAI.TaskPane.Variants
{
    /// <summary>
    /// Builds and drives the Variants tab. Generates 1-5 drafting variants
    /// from a single intent prompt + optional reasoning override, parses the
    /// fenced-JSON envelope via <see cref="VariantParser"/>, stores results
    /// in <see cref="VariantStore"/>, and renders card UI per variant.
    /// </summary>
    public sealed class VariantsController : IDisposable
    {
        private readonly Control _host;
        private readonly CodexChatService _chat;
        private readonly IToolHost _toolHost;
        private readonly LiveOutlookSurface _surface;
        private readonly Func<string, bool> _insertCallback;   // body -> success
        private readonly Func<string, bool> _replaceCallback;  // body -> success
        private readonly VariantStore _store = new VariantStore();
        private readonly VariantParser _parser = new VariantParser();

        // UI handles
        private TextBox _txtIntent;
        private NumericUpDown _numCount;
        private ComboBox _cmbReasoning;
        private Button _btnGenerate;
        private Button _btnRegenerateAll;
        private Button _btnCancel;
        private Label _lblStatus;
        private FlowLayoutPanel _cardsPanel;

        private CancellationTokenSource _activeCts;
        private bool _generating;
        private bool _isDisposed;

        public VariantsController(
            Control host,
            CodexChatService chat,
            IToolHost toolHost,
            LiveOutlookSurface surface,
            Func<string, bool> insertCallback,
            Func<string, bool> replaceCallback)
        {
            _host = host ?? throw new ArgumentNullException(nameof(host));
            _chat = chat ?? throw new ArgumentNullException(nameof(chat));
            _toolHost = toolHost ?? throw new ArgumentNullException(nameof(toolHost));
            _surface = surface;
            _insertCallback = insertCallback;
            _replaceCallback = replaceCallback;
        }

        public void BuildUi()
        {
            _host.Controls.Clear();
            _host.SuspendLayout();

            // Composer row: intent input + count + reasoning + Generate.
            var lblIntent = new Label
            {
                Text = "Drafting intent:",
                Font = new Font("Segoe UI", 9F),
                Location = new Point(8, 8),
                AutoSize = true
            };
            _txtIntent = new TextBox
            {
                Location = new Point(8, 28),
                Size = new Size(290, 50),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 9F),
                // (net472 has no PlaceholderText; tooltip approximates it.)
            };
            var intentTip = new ToolTip();
            intentTip.SetToolTip(_txtIntent,
                "What do you want to say? (e.g. 'decline politely but leave the door open for Q3')");

            var lblCount = new Label
            {
                Text = "Count:",
                Location = new Point(8, 86),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F)
            };
            _numCount = new NumericUpDown
            {
                Location = new Point(52, 84),
                Width = 50,
                Minimum = 1,
                Maximum = 5,
                Value = 3
            };

            var lblReasoning = new Label
            {
                Text = "Effort:",
                Location = new Point(112, 86),
                AutoSize = true,
                Font = new Font("Segoe UI", 9F)
            };
            _cmbReasoning = new ComboBox
            {
                Location = new Point(155, 84),
                Width = 70,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F)
            };
            _cmbReasoning.Items.AddRange(new object[] { "", "Minimal", "Low", "Medium", "High" });
            _cmbReasoning.SelectedIndex = 0;

            _btnGenerate = new Button
            {
                Text = "Generate",
                Location = new Point(228, 84),
                Width = 70,
                Height = 24,
                Font = new Font("Segoe UI", 9F)
            };
            _btnGenerate.Click += OnGenerateClick;

            _btnRegenerateAll = new Button
            {
                Text = "Regenerate all",
                Location = new Point(8, 114),
                Width = 110,
                Height = 22,
                Font = new Font("Segoe UI", 8F),
                Enabled = false
            };
            _btnRegenerateAll.Click += (s, e) => _ = GenerateAsync(replace: true);

            _btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(228, 114),
                Width = 70,
                Height = 22,
                Font = new Font("Segoe UI", 8F),
                Visible = false
            };
            _btnCancel.Click += (s, e) => { try { _activeCts?.Cancel(); } catch { } };

            _lblStatus = new Label
            {
                Location = new Point(8, 140),
                Size = new Size(290, 18),
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.DarkSlateGray
            };

            _cardsPanel = new FlowLayoutPanel
            {
                Location = new Point(8, 162),
                Size = new Size(295, 380),
                AutoScroll = true,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            _host.Controls.Add(lblIntent);
            _host.Controls.Add(_txtIntent);
            _host.Controls.Add(lblCount);
            _host.Controls.Add(_numCount);
            _host.Controls.Add(lblReasoning);
            _host.Controls.Add(_cmbReasoning);
            _host.Controls.Add(_btnGenerate);
            _host.Controls.Add(_btnRegenerateAll);
            _host.Controls.Add(_btnCancel);
            _host.Controls.Add(_lblStatus);
            _host.Controls.Add(_cardsPanel);

            _host.ResumeLayout();
        }

        private void OnGenerateClick(object sender, EventArgs e)
        {
            _ = GenerateAsync(replace: true);
        }

        public async Task GenerateAsync(bool replace, Tone? singleVariantTone = null)
        {
            if (_generating) return;
            var intent = (_txtIntent.Text ?? "").Trim();
            if (string.IsNullOrEmpty(intent))
            {
                SetStatus("Type a drafting intent first.", isError: true);
                return;
            }

            _generating = true;
            _activeCts = new CancellationTokenSource();
            SetUiEnabled(false);
            _btnCancel.Visible = true;
            SetStatus("Generating...", isError: false);

            try
            {
                int count = singleVariantTone.HasValue ? 1 : (int)_numCount.Value;
                var ctx = new ConversationContext
                {
                    SystemInstructions = BuildVariantsSystemPrompt(count, singleVariantTone),
                    // IMPORTANT: write tools off (no create_draft from Variants).
                    IncludeWriteTools = false,
                    ReasoningEffortOverride = NormalizeEffort(_cmbReasoning.SelectedItem as string)
                };
                var userMessage = BuildVariantsUserMessage(intent, count, singleVariantTone);
                var sink = new ChatEventSink();
                var result = await _chat.RunTurnAsync(ctx, userMessage, _toolHost, sink, _activeCts.Token);

                if (result.StopReason == StopReason.Cancelled)
                {
                    SetStatus("Cancelled.", false);
                    return;
                }
                if (result.StopReason == StopReason.Error)
                {
                    SetStatus("Error: " + (result.ErrorMessage ?? "unknown"), true);
                    return;
                }

                var variants = _parser.Parse(result.FinalAssistantText);
                if (variants.Count == 0)
                {
                    SetStatus("Model didn't return parseable variants. Try again or refine the intent.", true);
                    return;
                }

                if (replace || _store.Count == 0)
                {
                    _store.Replace(variants);
                }
                else
                {
                    // Single-variant regenerate path: caller passes singleVariantTone
                    // and we look for an existing entry with that tone to replace.
                    if (singleVariantTone.HasValue)
                    {
                        var snap = _store.Snapshot();
                        int idx = -1;
                        for (int i = 0; i < snap.Count; i++)
                        {
                            if (snap[i].Tone == singleVariantTone.Value) { idx = i; break; }
                        }
                        if (idx >= 0)
                        {
                            _store.Update(idx, variants[0]);
                        }
                        else
                        {
                            _store.Replace(variants);
                        }
                    }
                }

                RenderCards();
                SetStatus("Done. " + _store.Count + " variant(s) ready.", false);
                _btnRegenerateAll.Enabled = true;
            }
            catch (OperationCanceledException)
            {
                SetStatus("Cancelled.", false);
            }
            catch (Exception ex)
            {
                SetStatus("Error: " + ex.Message, true);
            }
            finally
            {
                _generating = false;
                SetUiEnabled(true);
                _btnCancel.Visible = false;
                _activeCts?.Dispose();
                _activeCts = null;
            }
        }

        private void RenderCards()
        {
            _cardsPanel.SuspendLayout();
            _cardsPanel.Controls.Clear();
            var snap = _store.Snapshot();
            for (int i = 0; i < snap.Count; i++)
            {
                _cardsPanel.Controls.Add(BuildVariantCard(i, snap[i]));
            }
            _cardsPanel.ResumeLayout();
        }

        private Panel BuildVariantCard(int index, Variant variant)
        {
            var card = new Panel
            {
                Size = new Size(275, 150),
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, 6),
                BackColor = Color.FromArgb(252, 251, 250)
            };

            var lblTone = new Label
            {
                Text = "  " + variant.Tone + "  ",
                Location = new Point(6, 6),
                AutoSize = true,
                BackColor = ToneColor(variant.Tone),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8F, FontStyle.Bold)
            };

            var lblChars = new Label
            {
                Text = (variant.Body ?? "").Length + " chars",
                Location = new Point(180, 6),
                AutoSize = true,
                Font = new Font("Segoe UI", 8F),
                ForeColor = Color.DarkSlateGray
            };

            var preview = FirstNLines(variant.Body ?? "", 3);
            var lblPreview = new Label
            {
                Text = preview,
                Location = new Point(6, 28),
                Size = new Size(263, 60),
                Font = new Font("Segoe UI", 9F),
                AutoEllipsis = true
            };

            var btnInsert = new Button
            {
                Text = "Insert",
                Location = new Point(6, 95),
                Width = 60,
                Height = 22,
                Font = new Font("Segoe UI", 8F)
            };
            btnInsert.Click += (s, e) =>
            {
                if (_insertCallback?.Invoke(variant.Body ?? "") ?? false)
                {
                    SetStatus("Inserted variant: " + variant.Tone, false);
                }
            };

            var btnReplace = new Button
            {
                Text = "Replace",
                Location = new Point(70, 95),
                Width = 60,
                Height = 22,
                Font = new Font("Segoe UI", 8F)
            };
            btnReplace.Click += (s, e) =>
            {
                if (_replaceCallback?.Invoke(variant.Body ?? "") ?? false)
                {
                    SetStatus("Replaced body with variant: " + variant.Tone, false);
                }
            };

            var btnRegen = new Button
            {
                Text = "Regenerate",
                Location = new Point(134, 95),
                Width = 80,
                Height = 22,
                Font = new Font("Segoe UI", 8F)
            };
            btnRegen.Click += (s, e) => _ = GenerateAsync(replace: false, singleVariantTone: variant.Tone);

            var lblRationale = new Label
            {
                Text = string.IsNullOrEmpty(variant.Rationale) ? "" : ("Rationale: " + variant.Rationale),
                Location = new Point(6, 122),
                Size = new Size(263, 22),
                Font = new Font("Segoe UI", 7.5F, FontStyle.Italic),
                ForeColor = Color.DimGray,
                AutoEllipsis = true
            };

            card.Controls.Add(lblTone);
            card.Controls.Add(lblChars);
            card.Controls.Add(lblPreview);
            card.Controls.Add(btnInsert);
            card.Controls.Add(btnReplace);
            card.Controls.Add(btnRegen);
            card.Controls.Add(lblRationale);
            return card;
        }

        private static string FirstNLines(string s, int n)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var lines = s.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            return string.Join("\r\n", lines.Take(n));
        }

        private static Color ToneColor(Tone tone)
        {
            switch (tone)
            {
                case Tone.Formal:       return Color.FromArgb(60, 60, 110);
                case Tone.Brief:        return Color.FromArgb(80, 100, 60);
                case Tone.Persuasive:   return Color.FromArgb(160, 80, 30);
                case Tone.Friendly:     return Color.FromArgb(200, 120, 50);
                case Tone.Technical:    return Color.FromArgb(50, 100, 130);
                case Tone.Apologetic:   return Color.FromArgb(140, 80, 80);
                case Tone.Direct:       return Color.FromArgb(80, 80, 80);
                case Tone.Diplomatic:   return Color.FromArgb(110, 90, 150);
                case Tone.Enthusiastic: return Color.FromArgb(180, 100, 80);
                default:                return Color.Gray;
            }
        }

        private static string BuildVariantsSystemPrompt(int count, Tone? specificTone)
        {
            var sb = new StringBuilder();
            sb.AppendLine("You are an AI email-drafting assistant generating multiple drafting variants.");
            sb.AppendLine();
            sb.Append("Produce ");
            if (specificTone.HasValue)
            {
                sb.Append("exactly 1 variant in the '").Append(specificTone.Value).Append("' tone.");
            }
            else
            {
                sb.Append("exactly ").Append(count).Append(" variants, each in a distinct tone.");
            }
            sb.AppendLine();
            sb.AppendLine();
            sb.AppendLine("Tone choices (pick from this set): Formal, Brief, Persuasive, Friendly, Technical, Apologetic, Direct, Diplomatic, Enthusiastic.");
            sb.AppendLine();
            sb.AppendLine("Hard constraints:");
            sb.AppendLine("- Do NOT call outlook_create_draft. You may use READ tools (search, read messages, list folders) for context, but writes are forbidden in this flow.");
            sb.AppendLine("- Each variant must include a non-empty `body` (the email text the user will paste).");
            sb.AppendLine("- Each variant should include a 1-sentence `rationale` explaining the tone choice.");
            sb.AppendLine("- Each variant should include a `subject` (may be omitted if the user is replying and the subject is fixed).");
            sb.AppendLine();
            sb.AppendLine("Output format: ONLY a single fenced JSON block, no prose outside the fence:");
            sb.AppendLine("```json");
            sb.AppendLine("{ \"variants\": [");
            sb.AppendLine("  { \"tone\": \"Formal\", \"rationale\": \"...\", \"subject\": \"...\", \"body\": \"...\" }");
            sb.AppendLine("]}");
            sb.AppendLine("```");
            return sb.ToString();
        }

        private static string BuildVariantsUserMessage(string intent, int count, Tone? specificTone)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Draft intent:");
            sb.AppendLine(intent);
            sb.AppendLine();
            if (specificTone.HasValue)
            {
                sb.AppendLine("Generate a single replacement variant in the '" + specificTone.Value + "' tone.");
            }
            else
            {
                sb.AppendLine("Generate " + count + " distinct variants spanning different tones.");
            }
            return sb.ToString();
        }

        private static string NormalizeEffort(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            return value;
        }

        private void SetUiEnabled(bool enabled)
        {
            _btnGenerate.Enabled = enabled;
            _btnRegenerateAll.Enabled = enabled && _store.Count > 0;
            _txtIntent.Enabled = enabled;
            _numCount.Enabled = enabled;
            _cmbReasoning.Enabled = enabled;
        }

        private void SetStatus(string text, bool isError)
        {
            if (_lblStatus.InvokeRequired)
            {
                _lblStatus.BeginInvoke(new Action(() => SetStatus(text, isError)));
                return;
            }
            _lblStatus.Text = text;
            _lblStatus.ForeColor = isError ? Color.DarkRed : Color.DarkSlateGray;
        }

        public void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            try { _activeCts?.Cancel(); } catch { }
        }
    }
}
