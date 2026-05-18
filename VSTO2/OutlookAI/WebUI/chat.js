/* ============================================================
   OutlookAI Chat - WebView2 surface controller (Phase 2)
   Public API (all on window.outlookai):
     appendUserMessage(text)
     appendAssistantMessage(messageId, initialText)
     appendTextDelta(messageId, delta)
     finalizeAssistantMessage(messageId, opts)
     appendToolCallCard(callId, name, argsJson)
     updateToolCallCard(callId, ok, summary, resultJson)
     appendAuditRow(text)
     showError(message)
     setComposerEnabled(enabled, isStopVisible)
     clear()
     applyTheme(themeName)              // "light" | "dark" | "high-contrast"
     setContextStrip({subject, recipients, thread})
   Host -> JS:
     C# calls these via WebView2.ExecuteScriptAsync("outlookai.X(...)").
   JS -> Host:
     window.chrome.webview.postMessage(JSON.stringify({type:..., payload:...}))
     Message types: 'send', 'stop', 'clear', 'copy', 'toolCardClicked'
   ============================================================ */

(function() {
  'use strict';

  // -- DOM refs ----------------------------------------------------
  var $messages = document.getElementById('messages');
  var $input = document.getElementById('composerInput');
  var $btnSend = document.getElementById('btnSend');
  var $btnStop = document.getElementById('btnStop');
  var $btnClear = document.getElementById('btnClear');
  var $btnCopy = document.getElementById('btnCopy');
  var $reasoning = document.getElementById('reasoningSelect');
  var $ctxSubject = document.getElementById('ctxSubject');
  var $ctxRecipients = document.getElementById('ctxRecipients');
  var $ctxThread = document.getElementById('ctxThread');
  var $quickActions = document.getElementById('quickActions');

  // -- Bridge to host ----------------------------------------------
  function postToHost(obj) {
    try {
      if (window.chrome && window.chrome.webview && window.chrome.webview.postMessage) {
        window.chrome.webview.postMessage(JSON.stringify(obj));
      } else {
        // Dev fallback: log to console.
        console.log('[hostPost]', obj);
      }
    } catch (err) {
      console.error('postToHost failed', err);
    }
  }

  // -- Minimal markdown renderer ----------------------------------
  // Supports: fenced code blocks (```), inline code, bold (**),
  // italic (*), links [text](url), unordered lists, ordered lists,
  // paragraphs. Anything else is escaped and rendered as plain text.
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function renderInline(line) {
    // Escape first, then re-introduce a small allow-list of markdown tags.
    var out = escapeHtml(line);
    // Inline code: `...` (must come before bold/italic).
    out = out.replace(/`([^`]+)`/g, function(m, code) {
      return '<code>' + code + '</code>';
    });
    // Bold: **...**
    out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    // Italic: *...*
    out = out.replace(/(^|[^*])\*([^*\n]+)\*([^*]|$)/g, '$1<em>$2</em>$3');
    // Links: [text](url)
    out = out.replace(/\[([^\]]+)\]\(([^)]+)\)/g, function(m, text, href) {
      // Only allow http/https/mailto.
      if (!/^(https?:|mailto:)/i.test(href)) return text;
      return '<a href="' + href + '" target="_blank" rel="noopener noreferrer">' + text + '</a>';
    });
    return out;
  }

  function renderMarkdown(src) {
    if (!src) return '';
    var html = [];
    var lines = src.split('\n');
    var inFence = false;
    var fenceBuf = [];
    var fenceLang = '';
    var listType = null; // 'ul' | 'ol'
    var listBuf = [];

    function flushList() {
      if (listType) {
        html.push('<' + listType + '>');
        for (var i = 0; i < listBuf.length; i++) {
          html.push('<li>' + renderInline(listBuf[i]) + '</li>');
        }
        html.push('</' + listType + '>');
        listType = null;
        listBuf = [];
      }
    }

    var paraBuf = [];
    function flushPara() {
      if (paraBuf.length > 0) {
        html.push('<p>' + renderInline(paraBuf.join(' ')) + '</p>');
        paraBuf = [];
      }
    }

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var fenceMatch = line.match(/^```(\w*)\s*$/);
      if (fenceMatch) {
        if (!inFence) {
          flushList(); flushPara();
          inFence = true;
          fenceLang = fenceMatch[1] || '';
          fenceBuf = [];
        } else {
          html.push('<pre><code' + (fenceLang ? ' class="lang-' + escapeHtml(fenceLang) + '"' : '') + '>' +
                    escapeHtml(fenceBuf.join('\n')) + '</code></pre>');
          inFence = false;
          fenceLang = '';
          fenceBuf = [];
        }
        continue;
      }
      if (inFence) {
        fenceBuf.push(line);
        continue;
      }
      // Lists
      var ulMatch = line.match(/^[-*]\s+(.+)$/);
      var olMatch = line.match(/^\d+\.\s+(.+)$/);
      if (ulMatch) {
        if (listType !== 'ul') { flushList(); flushPara(); listType = 'ul'; }
        listBuf.push(ulMatch[1]);
        continue;
      }
      if (olMatch) {
        if (listType !== 'ol') { flushList(); flushPara(); listType = 'ol'; }
        listBuf.push(olMatch[1]);
        continue;
      }
      // Paragraph break
      if (/^\s*$/.test(line)) {
        flushList(); flushPara();
        continue;
      }
      paraBuf.push(line);
    }
    if (inFence) {
      // Unterminated fence - treat the rest as plain text.
      html.push('<pre><code>' + escapeHtml(fenceBuf.join('\n')) + '</code></pre>');
    }
    flushList();
    flushPara();
    return html.join('\n');
  }

  // -- DOM-builder helpers ----------------------------------------
  function elt(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined && text !== null) e.textContent = text;
    return e;
  }

  function scrollToBottom() {
    $messages.scrollTop = $messages.scrollHeight;
  }

  var assistantMessages = {}; // id -> { container, content, raw }
  var toolCards = {};         // callId -> element

  // -- Public API --------------------------------------------------
  var api = {
    appendUserMessage: function(text) {
      var node = elt('div', 'msg msg-user');
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(text);
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    appendAssistantMessage: function(id, initialText) {
      var node = elt('div', 'msg msg-assistant');
      node.dataset.messageId = id;
      var content = elt('div', 'msg-content');
      content.innerHTML = renderMarkdown(initialText || '');
      node.appendChild(content);
      $messages.appendChild(node);
      assistantMessages[id] = {
        container: node,
        content: content,
        raw: initialText || ''
      };
      scrollToBottom();
    },

    appendTextDelta: function(id, delta) {
      var entry = assistantMessages[id];
      if (!entry) {
        api.appendAssistantMessage(id, delta);
        return;
      }
      entry.raw += delta;
      // For streaming, render the running text as markdown. This is
      // cheap enough for typical email-length replies.
      entry.content.innerHTML = renderMarkdown(entry.raw);
      scrollToBottom();
    },

    finalizeAssistantMessage: function(id, opts) {
      var entry = assistantMessages[id];
      if (!entry) return;
      opts = opts || {};
      if (opts.stopped) entry.container.classList.add('msg-stopped');
      if (opts.error) entry.container.classList.add('msg-error');
      scrollToBottom();
    },

    // Compact "working on it" status line. Replaces the previous verbose
    // expandable tool cards (the chat got too cluttered - user feedback).
    // Each tool call shows as a single italic line like
    //   ... Searching messages
    //   ... Reading message
    //   ... Listing folders
    // When the tool completes, the line is REMOVED entirely (we don't
    // need a permanent record - the model uses the result to produce
    // the actual assistant text, which is what the user reads).
    //
    // Friendly verbs per tool name. Anything not in the map falls back
    // to "Working on it..." so unknown future tools still render fine.
    appendToolCallCard: function(callId, name, argsJson) {
      var verb = ({
        outlook_get_current_compose_state: 'Reading compose context',
        outlook_get_current_selection:     'Reading current selection',
        outlook_list_folders:              'Listing folders',
        outlook_search_messages:           'Searching messages',
        outlook_read_message:              'Reading message',
        outlook_count_messages:            'Counting messages',
        outlook_list_recent_threads_with:  'Listing recent threads',
        outlook_create_draft:              'Creating draft',
        outlook_mark_as_read:              'Marking as read',
        outlook_flag_message:              'Flagging message',
        outlook_set_category:              'Setting category',
      })[name] || 'Working on it';

      var row = elt('div', 'tool-status');
      row.dataset.callId = callId;
      row.dataset.verb = verb;
      row.dataset.startedAt = String(Date.now());
      row.textContent = '\u2026 ' + verb;
      $messages.appendChild(row);

      // Live time counter. Updates the status line every second with the
      // elapsed seconds so the user can see the tool is making progress
      // rather than guessing whether Outlook has frozen. Counter math
      // uses Date.now() - startedAt so the displayed value is always
      // wall-clock-correct even if setInterval is throttled when the
      // WebView2 loses focus.
      var tick = function() {
        if (!row.parentNode) { return; }
        var ms = Date.now() - parseInt(row.dataset.startedAt, 10);
        var secs = Math.max(0, Math.floor(ms / 1000));
        if (secs > 0) {
          row.textContent = '\u2026 ' + verb + ' (' + secs + 's)';
        }
      };
      var tickId = setInterval(function() {
        if (!row.parentNode) { clearInterval(tickId); return; }
        tick();
      }, 1000);
      row.dataset.tickId = String(tickId);

      // Force an immediate refresh when the WebView2 regains focus.
      // WebView2 throttles setInterval while hidden; without this, the
      // counter appears to lag (or to "reset" perceptually) until the
      // next tick after focus returns.
      var visHandler = function() {
        if (document.visibilityState === 'visible') tick();
      };
      document.addEventListener('visibilitychange', visHandler);
      window.addEventListener('focus', visHandler);
      row.dataset.hasVisHandler = '1';
      row._oai_visHandler = visHandler;

      toolCards[callId] = row;
      scrollToBottom();
    },

    updateToolCallCard: function(callId, ok, summary, resultJson) {
      var row = toolCards[callId];
      if (!row) return;
      // Stop the live time counter.
      var tickId = parseInt(row.dataset.tickId || '0', 10);
      if (tickId) clearInterval(tickId);
      // Drop the visibility-change handler (one was attached per row).
      if (row._oai_visHandler) {
        try {
          document.removeEventListener('visibilitychange', row._oai_visHandler);
          window.removeEventListener('focus', row._oai_visHandler);
        } catch (e) { /* best-effort */ }
        row._oai_visHandler = null;
      }
      // On completion, drop the status line. Errors stick around as a
      // muted single-line error so the user sees that something failed
      // without the full JSON dump.
      if (ok) {
        if (row.parentNode) row.parentNode.removeChild(row);
      } else {
        row.classList.add('tool-status-err');
        row.textContent = '\u26A0 ' + (summary || 'tool error');
      }
      delete toolCards[callId];
      scrollToBottom();
    },

    appendAuditRow: function(text) {
      var row = elt('div', 'audit-row');
      var time = elt('span', 'audit-time', new Date().toLocaleTimeString());
      row.appendChild(time);
      row.appendChild(document.createTextNode(text));
      $messages.appendChild(row);
      scrollToBottom();
    },

    showError: function(message) {
      var node = elt('div', 'msg msg-assistant msg-error');
      var content = elt('div', 'msg-content');
      content.textContent = message || 'An error occurred.';
      node.appendChild(content);
      $messages.appendChild(node);
      scrollToBottom();
    },

    setComposerEnabled: function(enabled, isStopVisible) {
      $input.disabled = !enabled;
      $btnSend.disabled = !enabled;
      $btnSend.hidden = !!isStopVisible;
      $btnStop.hidden = !isStopVisible;
      $btnClear.disabled = !enabled && !isStopVisible;
      $btnCopy.disabled = false; // always allow copy
      if (enabled) $input.focus();
    },

    clear: function() {
      // Stop any live tool-status time counters and unbind any
      // visibility-change handlers before tossing their DOM.
      for (var cid in toolCards) {
        try {
          var row = toolCards[cid];
          var tickId = row && parseInt(row.dataset && row.dataset.tickId || '0', 10);
          if (tickId) clearInterval(tickId);
          if (row && row._oai_visHandler) {
            try {
              document.removeEventListener('visibilitychange', row._oai_visHandler);
              window.removeEventListener('focus', row._oai_visHandler);
            } catch (e) { /* best-effort */ }
            row._oai_visHandler = null;
          }
        } catch (e) { /* best-effort */ }
      }
      $messages.innerHTML = '';
      assistantMessages = {};
      toolCards = {};
    },

    applyTheme: function(themeName) {
      var allowed = ['light', 'dark', 'high-contrast'];
      var theme = allowed.indexOf(themeName) >= 0 ? themeName : 'light';
      document.body.classList.remove('theme-light', 'theme-dark', 'theme-high-contrast');
      document.body.classList.add('theme-' + theme);
    },

    setContextStrip: function(ctx) {
      ctx = ctx || {};
      // Phase 3a: support two shapes of context.
      //   Inbox shape:    { folder, unread_count, total_count, selection? }
      //   Compose shape:  { subject, recipients, thread }
      // Disambiguate by checking for ctx.folder.
      if (ctx.folder !== undefined) {
        var unread = (ctx.unread_count != null) ? (' (' + ctx.unread_count + ' unread)') : '';
        $ctxSubject.textContent = 'In: ' + ctx.folder + unread;
        if (ctx.selection && ctx.selection.count > 0) {
          if (ctx.selection.count === 1) {
            $ctxRecipients.textContent = 'Selected: ' + (ctx.selection.subject || '') +
              (ctx.selection.from ? (' \u2014 ' + ctx.selection.from) : '');
          } else {
            $ctxRecipients.textContent = 'Selected: ' + ctx.selection.count + ' messages';
          }
        } else {
          $ctxRecipients.textContent = '';
        }
        $ctxThread.textContent = '';
        return;
      }
      // Compose shape (Phase 2 behaviour unchanged).
      $ctxSubject.textContent = ctx.subject ? ('Re: ' + ctx.subject) : 'New email';
      var recipients = (ctx.recipients || []).join(', ');
      $ctxRecipients.textContent = recipients ? ('To: ' + recipients) : '';
      $ctxThread.textContent = ctx.thread || '';
    },

    /**
     * Render the row of quick-action chips above the composer. Each chip
     * is { label, prompt }. Clicking a chip pre-fills the textarea with
     * the prompt AND immediately sends (auto-send per the Phase 3a spec).
     * Calling with [] empties the row.
     */
    setQuickActions: function(chips) {
      if (!$quickActions) return;
      while ($quickActions.firstChild) $quickActions.removeChild($quickActions.firstChild);
      (chips || []).forEach(function(chip) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'qa-chip';
        btn.textContent = chip.label;
        btn.title = chip.prompt;
        btn.addEventListener('click', function() {
          $input.value = chip.prompt;
          sendInput();
        });
        $quickActions.appendChild(btn);
      });
    },

    /**
     * Populate the reasoning-effort dropdown based on the model the host
     * is currently configured for. C# computes the list via
     * Config.ReasoningEffortsForModel and pushes it here. Always keeps
     * a leading "(default)" option mapped to value="".
     *
     *  opts   - array of strings, e.g. ['None','Low','Medium','High','XHigh']
     *  selected - optional. Pre-selects the matching option (case-insensitive).
     *             Pass '' to keep the default.
     */
    setReasoningOptions: function(opts, selected) {
      while ($reasoning.firstChild) $reasoning.removeChild($reasoning.firstChild);
      var def = document.createElement('option');
      def.value = '';
      def.textContent = '(default)';
      $reasoning.appendChild(def);
      (opts || []).forEach(function(name) {
        var el = document.createElement('option');
        el.value = name;
        el.textContent = name;
        $reasoning.appendChild(el);
      });
      if (selected) {
        for (var i = 0; i < $reasoning.options.length; i++) {
          if ($reasoning.options[i].value.toLowerCase() === String(selected).toLowerCase()) {
            $reasoning.selectedIndex = i;
            break;
          }
        }
      }
    },

    // Used by C# during dev to verify the bridge is live before
    // pushing real messages. Returns a string from JS that C# can
    // assert on with await ExecuteScriptAsync("outlookai.ping()").
    ping: function() { return 'pong'; }
  };

  window.outlookai = api;

  // -- Composer event wiring --------------------------------------
  function sendInput() {
    // Guard against double-send when a previous turn is still running.
    // The host disables $input + $btnSend when a turn is in flight, but
    // a fast Enter or a race between WebMessageReceived and the C# side
    // re-enabling the composer can still let a second submission through
    // and produce two concurrent tool-status rows.
    if ($input.disabled || $btnSend.disabled) return;
    var text = $input.value.trim();
    if (!text) return;
    $input.value = '';
    postToHost({
      type: 'send',
      payload: {
        text: text,
        reasoning: $reasoning.value || null
      }
    });
  }

  $btnSend.addEventListener('click', sendInput);
  $btnStop.addEventListener('click', function() {
    postToHost({ type: 'stop' });
  });
  $btnClear.addEventListener('click', function() {
    postToHost({ type: 'clear' });
  });
  $btnCopy.addEventListener('click', function() {
    postToHost({ type: 'copy' });
  });

  // Enter sends, Shift+Enter inserts a newline (standard chat UX).
  $input.addEventListener('keydown', function(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendInput();
    }
  });

  // Tell the host we're ready so it can push the initial context strip
  // + apply the theme before the first user message.
  postToHost({ type: 'ready' });
})();
