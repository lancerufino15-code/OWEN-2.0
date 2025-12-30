/**
 * Markdown table parsing/rendering helpers used by the chat UI.
 *
 * Used by: `public/chat.js` and tests to detect and render tables from text.
 *
 * Key exports:
 * - Parsers for markdown/piped tables.
 * - Renderers for HTML and markdown output.
 *
 * Assumptions:
 * - Input text uses pipe-delimited markdown conventions; alignment rows are optional.
 */
const PIPE_MIN_COLUMNS = 2;
const ESCAPED_PIPE_TOKEN = "__OWEN_ESCAPED_PIPE__";
const ESCAPED_PIPE_RE = /\\\|/g;
const ESCAPED_PIPE_TOKEN_RE = new RegExp(ESCAPED_PIPE_TOKEN, "g");

function protectEscapedPipes(value) {
  return (value || "").replace(ESCAPED_PIPE_RE, ESCAPED_PIPE_TOKEN);
}

function restoreEscapedPipes(value) {
  return (value || "").replace(ESCAPED_PIPE_TOKEN_RE, "|");
}

/**
 * Normalize a line by unescaping leading pipe characters.
 *
 * @param line - Raw line string.
 * @returns Normalized line string.
 */
export function normalizeTableLine(line) {
  return (line || "").replace(/^\s*\\\|/, "|");
}

/**
 * Split a pipe-delimited row into trimmed cell values.
 *
 * @param line - Raw line string.
 * @returns Array of cell strings (empty when not a table row).
 */
export function splitPipeRow(line) {
  const normalized = protectEscapedPipes(normalizeTableLine(line));
  const parts = normalized.split("|");
  if (parts.length < 2) return [];
  if ((parts[0] || "").trim() === "") parts.shift();
  if ((parts[parts.length - 1] || "").trim() === "") parts.pop();
  return parts.map(part => restoreEscapedPipes(part).trim());
}

/**
 * Check whether a line is a markdown table divider row.
 *
 * @param line - Raw line string.
 * @returns True when the line matches divider syntax.
 */
export function isMarkdownTableDivider(line) {
  if (!line) return false;
  const trimmed = normalizeTableLine(line).trim();
  if (!trimmed.includes("-")) return false;
  return /^\|?\s*:?-{2,}.*\|.*$/.test(trimmed);
}

/**
 * Check whether a line looks like a markdown table row.
 *
 * @param line - Raw line string.
 * @returns True when the line has at least two columns.
 */
export function isMarkdownTableRow(line) {
  return splitPipeRow(line).length >= PIPE_MIN_COLUMNS;
}

/**
 * Check whether a line starts with a pipe and has table columns.
 *
 * @param line - Raw line string.
 * @returns True when the line looks like a pipe table row.
 */
export function isPipeTableRow(line) {
  if (!/^\s*\|/.test(line || "")) return false;
  return splitPipeRow(line).length >= PIPE_MIN_COLUMNS;
}

function getStableColumnCount(rows) {
  if (!rows || rows.length < 2) return null;
  const counts = rows.map(row => row.length);
  const freq = new Map();
  counts.forEach(count => freq.set(count, (freq.get(count) || 0) + 1));
  let bestCount = null;
  let bestFreq = 0;
  freq.forEach((value, key) => {
    if (value > bestFreq || (value === bestFreq && key > (bestCount || 0))) {
      bestFreq = value;
      bestCount = key;
    }
  });
  if (!bestCount || bestCount < PIPE_MIN_COLUMNS) return null;
  if (bestFreq < 2) return null;
  if (bestFreq / rows.length < 0.7) return null;
  return bestCount;
}

function inferPipeTableHeaders(rows, columnCount) {
  if (columnCount !== 2 || rows.length < 2) return null;
  const firstCells = rows.map(row => (row[0] || "").trim());
  const stepMatches = firstCells.filter(cell => /^(step|phase|stage|day|week|month)\s*\d+/i.test(cell)).length;
  if (stepMatches / rows.length >= 0.7) return ["Step", "Details"];
  return null;
}

function normalizeRowsToColumnCount(rows, columnCount) {
  return rows.map(row => {
    const normalized = row.slice(0, columnCount);
    while (normalized.length < columnCount) normalized.push("");
    return normalized;
  });
}

/**
 * Parse a markdown table block into headers, rows, and trailing text.
 *
 * @param block - Markdown block text.
 * @returns Parsed table object or null when not a table.
 */
export function parseMarkdownTableBlock(block) {
  if (!block || !block.includes("|")) return null;
  if (/```/.test(block)) return null;
  const lines = block
    .split(/\n/)
    .map(line => normalizeTableLine(line).trimEnd());
  const nonEmpty = lines.filter(line => line.trim());
  if (nonEmpty.length < 2) return null;

  let headerIdx = -1;
  for (let i = 0; i < lines.length - 1; i += 1) {
    if (isMarkdownTableRow(lines[i]) && isMarkdownTableDivider(lines[i + 1])) {
      headerIdx = i;
      break;
    }
  }

  if (headerIdx !== -1) {
    const caption = lines.slice(0, headerIdx).filter(line => line.trim()).join(" ");
    const headers = splitPipeRow(lines[headerIdx]);
    if (headers.length < PIPE_MIN_COLUMNS) return null;
    const bodyLines = [];
    let restIndex = lines.length;
    for (let i = headerIdx + 2; i < lines.length; i += 1) {
      const line = lines[i];
      if (isMarkdownTableRow(line)) {
        bodyLines.push(line);
        continue;
      }
      if (!line.trim()) {
        if (bodyLines.length) {
          restIndex = i;
          break;
        }
        continue;
      }
      restIndex = i;
      break;
    }
    if (!bodyLines.length) return null;
    const rows = normalizeRowsToColumnCount(bodyLines.map(splitPipeRow), headers.length);
    const trailingText = restIndex < lines.length ? lines.slice(restIndex).join("\n").trim() : "";
    return {
      caption,
      headers,
      rows,
      hasHeader: true,
      columnCount: headers.length,
      layout: headers.length === 2 ? "kv" : "grid",
      trailingText,
    };
  }

  const firstRowIndex = lines.findIndex(line => isMarkdownTableRow(line));
  if (firstRowIndex === -1) return null;
  const captionLines = lines.slice(0, firstRowIndex).filter(line => line.trim());
  const rowLines = lines.slice(firstRowIndex).filter(line => line.trim());
  if (rowLines.length < 2) return null;
  if (!rowLines.every(line => isMarkdownTableRow(line))) return null;

  const rows = rowLines.map(splitPipeRow);
  const columnCount = getStableColumnCount(rows);
  if (!columnCount) return null;

  const normalizedRows = normalizeRowsToColumnCount(rows, columnCount);
  const inferredHeaders = inferPipeTableHeaders(normalizedRows, columnCount);
  return {
    caption: captionLines.join(" "),
    headers: inferredHeaders || [],
    rows: normalizedRows,
    hasHeader: Boolean(inferredHeaders),
    columnCount,
    layout: columnCount === 2 ? "kv" : "grid",
  };
}

function sanitizeSpanValue(value) {
  const parsed = Number.parseInt(value || "", 10);
  if (!Number.isFinite(parsed) || parsed <= 1) return null;
  return parsed;
}

function countRowColumns(row) {
  if (!row) return 0;
  const cells = Array.from(row.children || []).filter(
    node => node && (node.tagName === "TD" || node.tagName === "TH"),
  );
  if (!cells.length) return 0;
  return cells.reduce((sum, cell) => sum + (sanitizeSpanValue(cell.getAttribute("colspan")) || 1), 0);
}

function extractCellText(cell) {
  if (!cell) return "";
  return (cell.textContent || "").replace(/\s+/g, " ").trim();
}

/**
 * Render a parsed table block to HTML.
 *
 * @param block - Parsed table block from `parseMarkdownTableBlock`.
 * @param opts - Render options (className, caption, etc.).
 * @returns HTML string for the table.
 */
export function renderHtmlTable(block, opts = {}) {
  if (!block || typeof document === "undefined") return null;
  if (!/<table[\s\S]*?>/i.test(block)) return null;
  const match = block.match(/<table[\s\S]*?<\/table>/i);
  if (!match) return null;
  const wrapper = document.createElement("div");
  wrapper.innerHTML = match[0];
  const rawTable = wrapper.querySelector("table");
  if (!rawTable) return null;
  const formatCell = typeof opts.formatCell === "function" ? opts.formatCell : null;
  const pickTone = typeof opts.pickTone === "function" ? opts.pickTone : null;
  const caption = extractCellText(rawTable.querySelector("caption"));

  const thead = rawTable.querySelector("thead");
  const tbody = rawTable.querySelector("tbody");
  const tfoot = rawTable.querySelector("tfoot");
  let headerRows = thead ? Array.from(thead.querySelectorAll("tr")) : [];
  let bodyRows = tbody ? Array.from(tbody.querySelectorAll("tr")) : [];
  const footerRows = tfoot ? Array.from(tfoot.querySelectorAll("tr")) : [];

  if (!tbody) {
    const allRows = Array.from(rawTable.querySelectorAll("tr")).filter(row => {
      if (thead && thead.contains(row)) return false;
      if (tfoot && tfoot.contains(row)) return false;
      return true;
    });
    bodyRows = allRows;
  }

  if (!headerRows.length && bodyRows.length) {
    const firstRow = bodyRows[0];
    if (firstRow && firstRow.querySelector("th")) {
      headerRows = [firstRow];
      bodyRows = bodyRows.slice(1);
    }
  }

  const columnCount = Math.max(
    0,
    ...[...headerRows, ...bodyRows, ...footerRows].map(countRowColumns),
  );
  const table = document.createElement("table");
  table.className = opts.tableClassName || "owen-table AnswerTable";
  table.dataset.answerTable = "true";
  if (columnCount) table.style.setProperty("--answer-table-columns", String(columnCount));
  if (columnCount === 2) table.classList.add("AnswerTable--kv");

  const headerTexts = [];
  const bodyTextRows = [];

  if (headerRows.length) {
    const theadOut = document.createElement("thead");
    headerRows.forEach(row => {
      const tr = document.createElement("tr");
      const cells = Array.from(row.children || []).filter(
        node => node && (node.tagName === "TD" || node.tagName === "TH"),
      );
      cells.forEach(cell => {
        const th = document.createElement("th");
        const text = extractCellText(cell);
        if (text) headerTexts.push(text);
        if (formatCell) {
          th.innerHTML = formatCell(text || " ");
        } else {
          th.textContent = text || " ";
        }
        const colSpan = sanitizeSpanValue(cell.getAttribute("colspan"));
        if (colSpan) th.colSpan = colSpan;
        const rowSpan = sanitizeSpanValue(cell.getAttribute("rowspan"));
        if (rowSpan) th.rowSpan = rowSpan;
        tr.appendChild(th);
      });
      theadOut.appendChild(tr);
    });
    if (theadOut.childNodes.length) table.appendChild(theadOut);
  }

  const tbodyOut = document.createElement("tbody");
  [...bodyRows, ...footerRows].forEach(row => {
    const tr = document.createElement("tr");
    const rowTexts = [];
    const cells = Array.from(row.children || []).filter(
      node => node && (node.tagName === "TD" || node.tagName === "TH"),
    );
    cells.forEach(cell => {
      const td = document.createElement("td");
      const text = extractCellText(cell);
      rowTexts.push(text);
      if (formatCell) {
        td.innerHTML = formatCell(text || " ");
      } else {
        td.textContent = text || " ";
      }
      const colSpan = sanitizeSpanValue(cell.getAttribute("colspan"));
      if (colSpan) td.colSpan = colSpan;
      const rowSpan = sanitizeSpanValue(cell.getAttribute("rowspan"));
      if (rowSpan) td.rowSpan = rowSpan;
      tr.appendChild(td);
    });
    if (rowTexts.length) bodyTextRows.push(rowTexts);
    tbodyOut.appendChild(tr);
  });
  if (tbodyOut.childNodes.length) table.appendChild(tbodyOut);

  if (pickTone) table.dataset.tone = pickTone(headerTexts, bodyTextRows);

  return {
    node: table,
    caption,
    headingHint: caption || headerTexts.join(" "),
  };
}

/**
 * Render a parsed table block back to markdown.
 *
 * @param block - Parsed table block from `parseMarkdownTableBlock`.
 * @param opts - Render options (caption, max widths, etc.).
 * @returns Markdown string.
 */
export function renderMarkdownTable(block, opts = {}) {
  const parsed = parseMarkdownTableBlock(block);
  if (!parsed) return null;
  if (typeof document === "undefined") return null;
  const formatCell = typeof opts.formatCell === "function" ? opts.formatCell : null;
  const pickTone = typeof opts.pickTone === "function" ? opts.pickTone : null;

  const table = document.createElement("table");
  table.className = opts.tableClassName || "owen-table AnswerTable";
  table.dataset.answerTable = "true";
  if (parsed.layout === "kv") table.classList.add("AnswerTable--kv");
  if (parsed.columnCount) table.style.setProperty("--answer-table-columns", String(parsed.columnCount));
  if (pickTone) table.dataset.tone = pickTone(parsed.headers, parsed.rows);

  if (parsed.hasHeader && parsed.headers.length) {
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    parsed.headers.forEach(cellText => {
      const th = document.createElement("th");
      if (formatCell) {
        th.innerHTML = formatCell(cellText);
      } else {
        th.textContent = cellText;
      }
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
  }

  const tbody = document.createElement("tbody");
  parsed.rows.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cellText => {
      const td = document.createElement("td");
      if (formatCell) {
        td.innerHTML = formatCell(cellText || " ");
      } else {
        td.textContent = cellText || " ";
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);

  return {
    node: table,
    caption: parsed.caption,
    headingHint: parsed.caption || parsed.headers.join(" "),
    trailingText: parsed.trailingText,
  };
}

/**
 * Heuristic check for whether a text block looks like a table.
 *
 * @param block - Raw text block.
 * @returns True when the block resembles a table.
 */
export function isTableLikeBlock(block) {
  if (!block || /```/.test(block)) return false;
  const lines = (block || "")
    .split(/\n/)
    .map(line => normalizeTableLine(line).trim())
    .filter(Boolean);
  if (lines.length < 2) return false;
  for (let i = 0; i < lines.length - 1; i += 1) {
    if (isMarkdownTableRow(lines[i]) && isMarkdownTableDivider(lines[i + 1])) return true;
  }
  const tableRows = lines.filter(isMarkdownTableRow);
  if (tableRows.length >= 2) return true;
  return false;
}
