/**
 * ============================================================================
 *  CXOne Intent Scraper — Browser Console Version
 * ============================================================================
 *
 *  HOW TO USE:
 *    1. Open CXOne in Chrome and navigate to the Intent Builder page
 *       (the page with the Category > Topic > Intent kanban tree).
 *    2. Open Chrome DevTools (F12 or Ctrl+Shift+I).
 *    3. Go to the "Console" tab.
 *    4. Copy-paste this ENTIRE script into the console and press Enter.
 *    5. The script will:
 *       - Expand every collapsed tree node
 *       - Click each intent to load its detail panel
 *       - Extract Category, Topic, Intent, Percentage, Volume, Examples, Active
 *       - Download the result as an Excel (.xlsx) file
 *
 *  NOTE: The script includes a built-in lightweight XLSX writer so it has
 *        zero external dependencies. Just paste and run.
 * ============================================================================
 */

(async function CXOneIntentScraper() {
  'use strict';

  // ── Configuration ──────────────────────────────────────────────────────
  const EXPAND_DELAY   = 1000;  // ms to wait after expanding a tree node
  const CLICK_DELAY    = 2000;  // ms to wait after clicking an intent for detail panel
  const BETWEEN_CLICKS = 300;   // ms between sequential intent clicks

  // ── Helpers ────────────────────────────────────────────────────────────
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function log(msg) {
    console.log(`%c[CXOne Scraper] ${msg}`, 'color: #2196F3; font-weight: bold;');
  }

  function logProgress(current, total, text) {
    const pct = ((current / total) * 100).toFixed(1);
    console.log(
      `%c[CXOne Scraper] [${current}/${total} — ${pct}%] ${text}`,
      'color: #4CAF50;'
    );
  }

  function logWarn(msg) {
    console.log(`%c[CXOne Scraper] ⚠ ${msg}`, 'color: #FF9800; font-weight: bold;');
  }

  // ── Step 1: Verify we are on the right page ────────────────────────────
  log('Starting CXOne Intent Scraper...');

  const kanbanPanel = document.querySelector('.kanban-view-panel');
  if (!kanbanPanel) {
    logWarn('Could not find .kanban-view-panel on this page.');
    logWarn('Make sure you are on the CXOne Intent Builder page.');
    return;
  }

  // ── Step 2: Expand all collapsed tree nodes ────────────────────────────
  log('Step 1/4: Expanding all collapsed tree nodes...');

  let expandedTotal = 0;
  let passNum = 0;

  while (true) {
    passNum++;
    // Collapsed nodes have a chevronrighticon inside the toggler button
    const collapsedTogglers = kanbanPanel.querySelectorAll(
      'p-treenode .p-tree-toggler:has(chevronrighticon)'
    );

    // Filter to only visible ones (not inside hidden panels)
    const visible = Array.from(collapsedTogglers).filter(
      (el) => el.offsetParent !== null
    );

    if (visible.length === 0) break;

    log(`  Pass ${passNum}: expanding ${visible.length} collapsed node(s)...`);

    for (const toggler of visible) {
      toggler.scrollIntoView({ block: 'center', behavior: 'instant' });
      toggler.click();
      await sleep(EXPAND_DELAY);
      expandedTotal++;
    }
  }

  log(`  Done — expanded ${expandedTotal} node(s) across ${passNum} pass(es).`);

  // ── Step 3: Collect tree hierarchy ─────────────────────────────────────
  log('Step 2/4: Collecting Category > Topic > Intent hierarchy...');

  const allNodes = kanbanPanel.querySelectorAll('.kanban-tree-node');
  const intentList = [];
  let currentCategory = '';
  let currentTopic = '';

  for (const node of allNodes) {
    const cls = node.className;
    const nameEl = node.querySelector('.kanban-tree-node-name');
    const pctEl = node.querySelector('.kanban-tree-node-statistics .percentage');
    const name = nameEl ? nameEl.textContent.trim() : '';
    const pct = pctEl ? pctEl.textContent.trim() : '0%';

    if (cls.includes('node-level-1')) {
      currentCategory = name;
      currentTopic = '';
    } else if (cls.includes('node-level-2')) {
      currentTopic = name;
    } else if (cls.includes('node-level-3')) {
      intentList.push({
        category: currentCategory,
        topic: currentTopic,
        intent: name,
        intentPercentage: pct,
        volume: '',
        examples: '',
        active: '',
        // Keep reference to the DOM node for clicking
        _nodeEl: node,
      });
    }
  }

  log(`  Found ${intentList.length} Level-3 intents.`);

  // ── Step 4: Click each intent and scrape detail panel ──────────────────
  log('Step 3/4: Clicking each intent to extract details...');
  log(`  This will take approximately ${Math.ceil((intentList.length * (CLICK_DELAY + BETWEEN_CLICKS)) / 1000 / 60)} minutes.`);

  for (let i = 0; i < intentList.length; i++) {
    const item = intentList[i];

    // Click the intent's p-treenode-content (the selectable row)
    const treeContent = item._nodeEl.closest('.p-treenode-content') || item._nodeEl;
    treeContent.scrollIntoView({ block: 'center', behavior: 'instant' });
    treeContent.click();

    await sleep(CLICK_DELAY);

    // ── Scrape the detail/info panel ──
    // Strategy: look for .info-body .info-item elements with .sub-title / .item-value
    const infoItems = document.querySelectorAll('.info-body .info-item');
    for (const infoItem of infoItems) {
      const titleEl = infoItem.querySelector('.sub-title');
      const valueEl = infoItem.querySelector('.item-value');
      if (!titleEl || !valueEl) continue;

      const title = titleEl.textContent.trim().toLowerCase();
      const value = valueEl.textContent.trim();

      if (title.includes('volume')) {
        item.volume = value;
      } else if (
        title.includes('example') ||
        title.includes('sample') ||
        title.includes('utterance') ||
        title.includes('training')
      ) {
        item.examples = value;
      } else if (title.includes('active') || title.includes('status')) {
        item.active = value;
      }
    }

    // Alternative: look for examples in a separate list/table
    if (!item.examples) {
      const exampleEls = document.querySelectorAll(
        '.sample-list .sample-item, ' +
        '.example-row, .intent-examples .example-text, ' +
        '.info-body .examples-list li, ' +
        '.sentence-list .sentence-item, ' +
        '.training-phrases .phrase-text'
      );
      if (exampleEls.length > 0) {
        item.examples = Array.from(exampleEls)
          .map((el) => el.textContent.trim())
          .filter(Boolean)
          .join('\n');
      }
    }

    // Alternative: look for volume near percentage
    if (!item.volume) {
      const statEls = document.querySelectorAll(
        '.panel-header .stat-value, .intent-stats .volume, ' +
        '.info-header .volume-value'
      );
      for (const stat of statEls) {
        const text = stat.textContent.trim();
        if (text && /\d/.test(text)) {
          item.volume = text;
          break;
        }
      }
    }

    // Check active status via toggle/checkbox
    if (!item.active) {
      const toggle = document.querySelector(
        '.active-toggle input[type="checkbox"], ' +
        '.info-body .active-status, ' +
        'p-checkbox[name*="active"] input'
      );
      if (toggle) {
        item.active = toggle.checked ? 'Yes' : 'No';
      }
    }

    logProgress(
      i + 1,
      intentList.length,
      `${item.category} > ${item.topic} > ${item.intent}`
    );

    await sleep(BETWEEN_CLICKS);
  }

  // Clean up DOM references before export
  const rows = intentList.map(({ _nodeEl, ...rest }) => rest);

  // ── Step 5: Generate and download Excel file ───────────────────────────
  log('Step 4/4: Generating Excel file...');

  downloadExcel(rows);

  // Print summary
  const categories = [...new Set(rows.map((r) => r.category))];
  const withExamples = rows.filter((r) => r.examples).length;
  log('');
  log('=== COMPLETE ===');
  log(`  Categories:      ${categories.length}`);
  log(`  Total intents:   ${rows.length}`);
  log(`  With examples:   ${withExamples}`);
  log(`  File downloaded: CXOne_Intents_Output.xlsx`);

  // Also log to a table for quick review
  console.table(rows.slice(0, 10));
  if (rows.length > 10) {
    log(`  ... and ${rows.length - 10} more rows (see the downloaded Excel file).`);
  }

  // ────────────────────────────────────────────────────────────────────────
  //  Minimal XLSX Generator (no external dependencies)
  //  Creates a valid .xlsx file using JSZip-free approach with raw ZIP.
  // ────────────────────────────────────────────────────────────────────────

  function downloadExcel(data) {
    const headers = [
      'Category',
      'Topic',
      'Intent',
      'Intent Percentage',
      'Volume',
      'Examples',
      'Active',
    ];
    const keys = [
      'category',
      'topic',
      'intent',
      'intentPercentage',
      'volume',
      'examples',
      'active',
    ];

    // Build CSV as a fallback-friendly format, then also produce XLSX
    // We'll generate a proper XLSX using XML + ZIP

    // ── Build sheet XML ──
    function escapeXml(s) {
      if (s == null) return '';
      return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
    }

    function colLetter(idx) {
      let s = '';
      idx++;
      while (idx > 0) {
        idx--;
        s = String.fromCharCode(65 + (idx % 26)) + s;
        idx = Math.floor(idx / 26);
      }
      return s;
    }

    let sheetRows = '';
    // Header row
    sheetRows += '<row r="1">';
    for (let c = 0; c < headers.length; c++) {
      const ref = colLetter(c) + '1';
      sheetRows += `<c r="${ref}" t="inlineStr" s="1"><is><t>${escapeXml(headers[c])}</t></is></c>`;
    }
    sheetRows += '</row>';

    // Data rows
    for (let r = 0; r < data.length; r++) {
      const rowNum = r + 2;
      sheetRows += `<row r="${rowNum}">`;
      for (let c = 0; c < keys.length; c++) {
        const ref = colLetter(c) + rowNum;
        const val = data[r][keys[c]] || '';
        sheetRows += `<c r="${ref}" t="inlineStr"><is><t>${escapeXml(val)}</t></is></c>`;
      }
      sheetRows += '</row>';
    }

    const lastCol = colLetter(headers.length - 1);
    const lastRow = data.length + 1;

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>
    <col min="1" max="1" width="25" customWidth="1"/>
    <col min="2" max="2" width="30" customWidth="1"/>
    <col min="3" max="3" width="45" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="12" customWidth="1"/>
    <col min="6" max="6" width="60" customWidth="1"/>
    <col min="7" max="7" width="10" customWidth="1"/>
  </cols>
  <sheetData>${sheetRows}</sheetData>
  <autoFilter ref="A1:${lastCol}${lastRow}"/>
</worksheet>`;

    const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
  </fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs>
  <cellXfs count="2">
    <xf/>
    <xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
</styleSheet>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Intents" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

    const workbookRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

    const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

    // ── Build ZIP (minimal implementation) ──
    const files = [
      { name: '[Content_Types].xml', content: contentTypesXml },
      { name: '_rels/.rels', content: relsXml },
      { name: 'xl/workbook.xml', content: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', content: workbookRelsXml },
      { name: 'xl/styles.xml', content: stylesXml },
      { name: 'xl/worksheets/sheet1.xml', content: sheetXml },
    ];

    const blob = buildZipBlob(files);

    // Trigger download
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'CXOne_Intents_Output.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    log('  Excel file download triggered.');

    // Also download CSV as backup
    downloadCsv(data, headers, keys);
  }

  function downloadCsv(data, headers, keys) {
    const csvRows = [headers.join(',')];
    for (const row of data) {
      csvRows.push(
        keys.map((k) => {
          const v = (row[k] || '').replace(/"/g, '""');
          return `"${v}"`;
        }).join(',')
      );
    }
    const csvBlob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
    const url = URL.createObjectURL(csvBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'CXOne_Intents_Output.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    log('  CSV backup also downloaded.');
  }

  // ── Minimal ZIP builder (STORE, no compression needed for XML) ─────────
  function buildZipBlob(files) {
    const enc = new TextEncoder();
    const parts = [];
    const centralDir = [];
    let offset = 0;

    for (const file of files) {
      const nameBytes = enc.encode(file.name);
      const contentBytes = enc.encode(file.content);

      const crc = crc32(contentBytes);
      const size = contentBytes.length;

      // Local file header (30 + nameLen + contentLen)
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const lv = new DataView(localHeader.buffer);
      lv.setUint32(0, 0x04034b50, true);  // signature
      lv.setUint16(4, 20, true);           // version needed
      lv.setUint16(6, 0, true);            // flags
      lv.setUint16(8, 0, true);            // compression: STORE
      lv.setUint16(10, 0, true);           // mod time
      lv.setUint16(12, 0, true);           // mod date
      lv.setUint32(14, crc, true);         // crc32
      lv.setUint32(18, size, true);        // compressed size
      lv.setUint32(22, size, true);        // uncompressed size
      lv.setUint16(26, nameBytes.length, true);
      lv.setUint16(28, 0, true);           // extra field length
      localHeader.set(nameBytes, 30);

      parts.push(localHeader, contentBytes);

      // Central directory entry
      const cdEntry = new Uint8Array(46 + nameBytes.length);
      const cv = new DataView(cdEntry.buffer);
      cv.setUint32(0, 0x02014b50, true);   // signature
      cv.setUint16(4, 20, true);           // version made by
      cv.setUint16(6, 20, true);           // version needed
      cv.setUint16(8, 0, true);            // flags
      cv.setUint16(10, 0, true);           // compression
      cv.setUint16(12, 0, true);           // mod time
      cv.setUint16(14, 0, true);           // mod date
      cv.setUint32(16, crc, true);
      cv.setUint32(20, size, true);
      cv.setUint32(24, size, true);
      cv.setUint16(28, nameBytes.length, true);
      cv.setUint16(30, 0, true);           // extra field length
      cv.setUint16(32, 0, true);           // comment length
      cv.setUint16(34, 0, true);           // disk number
      cv.setUint16(36, 0, true);           // internal attributes
      cv.setUint32(38, 0, true);           // external attributes
      cv.setUint32(42, offset, true);      // local header offset
      cdEntry.set(nameBytes, 46);

      centralDir.push(cdEntry);
      offset += localHeader.length + contentBytes.length;
    }

    const cdOffset = offset;
    let cdSize = 0;
    for (const cd of centralDir) {
      parts.push(cd);
      cdSize += cd.length;
    }

    // End of central directory
    const eocd = new Uint8Array(22);
    const ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(4, 0, true);
    ev.setUint16(6, 0, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, cdOffset, true);
    ev.setUint16(20, 0, true);
    parts.push(eocd);

    return new Blob(parts, {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc ^= bytes[i];
      for (let j = 0; j < 8; j++) {
        crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
      }
    }
    return (crc ^ 0xffffffff) >>> 0;
  }
})();
