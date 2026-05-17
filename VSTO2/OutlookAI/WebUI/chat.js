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

    appendToolCallCard: function(callId, name, argsJson) {
      var card = elt('div', 'tool-card tool-pending');
      card.dataset.callId = callId;
      var header = elt('div', 'tool-card-header');
      var glyph = elt('span', 'tool-card-glyph', '\u2022');
      var nameEl = elt('span', 'tool-card-name', name);
      var summary = elt('span', 'tool-card-summary', 'running...');
      var toggle = elt('span', 'tool-card-toggle', '[+]');
      header.appendChild(glyph);
      header.appendChild(nameEl);
      header.appendChild(summary);
      header.appendChild(toggle);
      var body = elt('div', 'tool-card-body');
      var argsLabel = elt('div', 'tool-card-label', 'arguments');
      var argsBox = elt('div', 'tool-card-args');
      argsBox.textContent = argsJson || '{}';
      body.appendChild(argsLabel);
      body.appendChild(argsBox);
      card.appendChild(header);
      card.appendChild(body);
      header.addEventListener('click', function() {
        var nowExpanded = !card.classList.contains('expanded');
        card.classList.toggle('expanded', nowExpanded);
        toggle.textContent = nowExpanded ? '[-]' : '[+]';
      });
      $messages.appendChild(card);
      toolCards[callId] = card;
      scrollToBottom();
    },

    updateToolCallCard: function(callId, ok, summary, resultJson) {
      var card = toolCards[callId];
      if (!card) return;
      card.classList.remove('tool-pending');
      card.classList.add(ok ? 'tool-ok' : 'tool-err');
      var glyph = card.querySelector('.tool-card-glyph');
      var summaryEl = card.querySelector('.tool-card-summary');
      var body = card.querySelector('.tool-card-body');
      glyph.textContent = ok ? '\u2713' : '\u26A0';
      summaryEl.textContent = summary || (ok ? 'ok' : 'error');
      var resultLabel = elt('div', 'tool-card-label', 'result');
      var resultBox = elt('div', 'tool-card-result');
      resultBox.textContent = resultJson || '{}';
      body.appendChild(resultLabel);
      body.appendChild(resultBox);
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
      $ctxSubject.textContent = ctx.subject ? ('Re: ' + ctx.subject) : 'New email';
      var recipients = (ctx.recipients || []).join(', ');
      $ctxRecipients.textContent = recipients ? ('To: ' + recipients) : '';
      $ctxThread.textContent = ctx.thread || '';
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
