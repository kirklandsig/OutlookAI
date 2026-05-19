/* ============================================================
   OutlookAI Markdown renderer
   ============================================================ */

(function(window) {
  'use strict';

  // Supports: ATX headers (#..######), fenced code blocks (```),
  // inline code, bold (**), italic (*), links [text](url), unordered
  // lists, ordered lists, GitHub-flavored tables (|col|col| with
  // |---|---| separator), horizontal rules (---), blockquotes (>),
  // and GitHub-style single-newline = <br> inside paragraphs so reports
  // with terse line-broken output render legibly.
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

  // Split a markdown table row into cells, trimming leading/trailing
  // pipes and stripping whitespace. Returns null for non-table rows.
  function splitTableRow(line) {
    if (!/^\s*\|.*\|\s*$/.test(line)) return null;
    var body = line.trim().replace(/^\|/, '').replace(/\|$/, '');
    return body.split('|').map(function (c) { return c.trim(); });
  }

  // Is this the separator row that follows the header in a GFM table?
  // Cells contain only - and : (with optional whitespace).
  function isTableSeparator(line) {
    var cells = splitTableRow(line);
    if (!cells || cells.length === 0) return false;
    for (var i = 0; i < cells.length; i++) {
      if (!/^:?-{1,}:?$/.test(cells[i])) return false;
    }
    return true;
  }

  function render(src) {
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
        // GFM-style: join paragraph lines with <br> so single-newline
        // text the model produced (terse digests, indented items
        // without bullet markers, address blocks, etc.) renders with
        // visible line breaks rather than collapsing onto one line.
        var rendered = paraBuf.map(renderInline).join('<br>');
        html.push('<p>' + rendered + '</p>');
        paraBuf = [];
      }
    }

    function flushAll() { flushList(); flushPara(); }

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];

      // Fenced code blocks first (highest precedence).
      var fenceMatch = line.match(/^```(\w*)\s*$/);
      if (fenceMatch) {
        if (!inFence) {
          flushAll();
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

      // ATX headers: # Title .. ###### Title
      var headerMatch = line.match(/^(#{1,6})\s+(.+?)\s*#*\s*$/);
      if (headerMatch) {
        flushAll();
        // Compact UI: clamp to h3..h6 so a top-level # does not produce
        // a giant heading inside a 340px-wide pane.
        var level = Math.min(6, Math.max(3, headerMatch[1].length + 2));
        html.push('<h' + level + '>' + renderInline(headerMatch[2]) + '</h' + level + '>');
        continue;
      }

      // Horizontal rule
      if (/^\s*([-*_])\s*\1\s*\1[\s\1]*$/.test(line)) {
        flushAll();
        html.push('<hr>');
        continue;
      }

      // GFM table: header row + separator row + body rows
      var headerCells = splitTableRow(line);
      if (headerCells && i + 1 < lines.length && isTableSeparator(lines[i + 1])) {
        flushAll();
        var tableHtml = ['<table class="md-table">'];
        tableHtml.push('<thead><tr>');
        headerCells.forEach(function (c) {
          tableHtml.push('<th>' + renderInline(c) + '</th>');
        });
        tableHtml.push('</tr></thead>');
        tableHtml.push('<tbody>');
        i += 2; // skip header + separator
        while (i < lines.length) {
          var rowCells = splitTableRow(lines[i]);
          if (!rowCells) break;
          tableHtml.push('<tr>');
          for (var j = 0; j < headerCells.length; j++) {
            tableHtml.push('<td>' + renderInline(rowCells[j] || '') + '</td>');
          }
          tableHtml.push('</tr>');
          i++;
        }
        tableHtml.push('</tbody></table>');
        html.push(tableHtml.join(''));
        i--; // outer loop will i++ past the last row
        continue;
      }

      // Blockquote
      var quoteMatch = line.match(/^>\s?(.*)$/);
      if (quoteMatch) {
        flushAll();
        html.push('<blockquote>' + renderInline(quoteMatch[1]) + '</blockquote>');
        continue;
      }

      // Lists
      var ulMatch = line.match(/^\s*[-*]\s+(.+)$/);
      var olMatch = line.match(/^\s*\d+\.\s+(.+)$/);
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

      // Paragraph break on blank line
      if (/^\s*$/.test(line)) {
        flushAll();
        continue;
      }

      paraBuf.push(line);
    }
    if (inFence) {
      html.push('<pre><code>' + escapeHtml(fenceBuf.join('\n')) + '</code></pre>');
    }
    flushAll();
    return html.join('\n');
  }

  window.markdown = window.markdown || {};
  window.markdown.render = render;
  window.markdown.escapeHtml = escapeHtml;
})(window);
