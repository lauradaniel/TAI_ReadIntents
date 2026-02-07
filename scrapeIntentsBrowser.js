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
 *       - Click the FIRST intent and discover the detail panel structure
 *       - Click each remaining intent and extract examples
 *       - Download the result as .xlsx and .csv files
 *
 *  NOTE: Zero external dependencies. Just paste and run.
 * ============================================================================
 */

(async function CXOneIntentScraper() {
  'use strict';

  // ── Configuration ──────────────────────────────────────────────────────
  const EXPAND_DELAY   = 1000;  // ms to wait after expanding a tree node
  const CLICK_DELAY    = 2500;  // ms to wait after clicking an intent
  const BETWEEN_CLICKS = 300;   // ms between sequential intent clicks

  // ── Helpers ────────────────────────────────────────────────────────────
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function log(msg) {
    console.log('%c[CXOne Scraper] ' + msg, 'color: #2196F3; font-weight: bold;');
  }
  function logProgress(cur, tot, text) {
    console.log('%c[CXOne Scraper] [' + cur + '/' + tot + ' - ' + ((cur/tot)*100).toFixed(1) + '%] ' + text, 'color: #4CAF50;');
  }
  function logWarn(msg) {
    console.log('%c[CXOne Scraper] WARNING: ' + msg, 'color: #FF9800; font-weight: bold;');
  }
  function logDebug(msg) {
    console.log('%c[CXOne Scraper][DEBUG] ' + msg, 'color: #9E9E9E;');
  }

  // ── Step 1: Verify page ────────────────────────────────────────────────
  log('Starting CXOne Intent Scraper...');

  const kanbanPanel = document.querySelector('.kanban-view-panel');
  if (!kanbanPanel) {
    logWarn('Could not find .kanban-view-panel. Are you on the Intent Builder page?');
    return;
  }

  // ── Step 2: Expand all collapsed tree nodes ────────────────────────────
  log('Step 1/5: Expanding all collapsed tree nodes...');

  let expandedTotal = 0;
  let passNum = 0;

  while (true) {
    passNum++;
    const collapsedTogglers = kanbanPanel.querySelectorAll(
      'p-treenode .p-tree-toggler:has(chevronrighticon)'
    );
    const visible = Array.from(collapsedTogglers).filter(el => el.offsetParent !== null);
    if (visible.length === 0) break;

    log('  Pass ' + passNum + ': expanding ' + visible.length + ' collapsed node(s)...');
    for (const toggler of visible) {
      toggler.scrollIntoView({ block: 'center', behavior: 'instant' });
      toggler.click();
      await sleep(EXPAND_DELAY);
      expandedTotal++;
    }
  }
  log('  Done - expanded ' + expandedTotal + ' node(s).');

  // ── Step 3: Collect tree hierarchy ─────────────────────────────────────
  log('Step 2/5: Collecting Category > Topic > Intent hierarchy...');

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
        _nodeEl: node,
      });
    }
  }

  log('  Found ' + intentList.length + ' Level-3 intents.');
  if (intentList.length === 0) {
    logWarn('No Level-3 intents found. Is the kanban tree fully expanded?');
    return;
  }

  // ── Step 4: DISCOVERY — snapshot DOM before/after clicking first intent ─
  log('Step 3/5: Discovery — clicking first intent to find the detail panel...');

  // Snapshot all visible leaf-text BEFORE clicking
  function snapshotLeafTexts() {
    const result = new Map();
    const all = document.querySelectorAll('*');
    for (const el of all) {
      if (el.children.length > 0) continue;
      if (el.offsetParent === null && el.offsetHeight === 0) continue;
      const txt = el.textContent.trim();
      if (txt.length < 5) continue;
      // Build a path to identify this element
      const path = buildPath(el);
      if (!result.has(path)) {
        result.set(path, txt);
      }
    }
    return result;
  }

  function buildPath(el) {
    const parts = [];
    let cur = el;
    let depth = 0;
    while (cur && cur !== document.body && depth < 10) {
      let s = cur.tagName.toLowerCase();
      if (cur.className && typeof cur.className === 'string') {
        const c = cur.className.trim().split(/\s+/).filter(x => !x.startsWith('ng-')).slice(0, 2).join('.');
        if (c) s += '.' + c;
      }
      parts.unshift(s);
      cur = cur.parentElement;
      depth++;
    }
    return parts.join(' > ');
  }

  const beforeSnap = snapshotLeafTexts();
  logDebug('Before-click snapshot: ' + beforeSnap.size + ' leaf text nodes');

  // Click the first intent
  const firstItem = intentList[0];
  const firstContent = firstItem._nodeEl.closest('.p-treenode-content') || firstItem._nodeEl;
  firstContent.scrollIntoView({ block: 'center', behavior: 'instant' });
  firstContent.click();
  await sleep(CLICK_DELAY + 1500); // Extra wait for first load

  const afterSnap = snapshotLeafTexts();
  logDebug('After-click snapshot: ' + afterSnap.size + ' leaf text nodes');

  // Find NEW elements that appeared after clicking
  const newTexts = [];
  for (const [path, txt] of afterSnap) {
    if (!beforeSnap.has(path)) {
      newTexts.push({ path, txt });
    }
  }

  log('  ' + newTexts.length + ' new text elements appeared after clicking.');

  // Log all new text for debugging
  logDebug('=== NEW text after clicking "' + firstItem.intent + '" ===');
  for (const item of newTexts) {
    logDebug('  [' + item.path + '] "' + item.txt.substring(0, 120) + '"');
  }

  // Find which new elements are likely examples (sentence-like, 15+ chars,
  // not percentages, not the intent name itself)
  const exampleCandidates = newTexts.filter(item => {
    const t = item.txt;
    if (t.length < 15 || t.length > 500) return false;
    if (/^\d+(\.\d+)?%?$/.test(t)) return false;
    if (t === firstItem.intent) return false;
    if (t === firstItem.category) return false;
    if (t === firstItem.topic) return false;
    return true;
  });

  logDebug('Example candidates: ' + exampleCandidates.length);
  for (const c of exampleCandidates) {
    logDebug('  EXAMPLE? "' + c.txt.substring(0, 100) + '" @ ' + c.path);
  }

  // Try to find a common parent selector pattern for the example elements
  let discoveredSelector = null;

  if (exampleCandidates.length > 0) {
    // Find the CSS class pattern shared by examples
    // Extract the most specific class from each path
    const classPatterns = {};
    for (const c of exampleCandidates) {
      // Get the last segment of the path (the actual element)
      const lastSeg = c.path.split(' > ').pop();
      classPatterns[lastSeg] = (classPatterns[lastSeg] || 0) + 1;
    }

    // Sort by count - the most repeated pattern is our selector
    const sorted = Object.entries(classPatterns).sort((a, b) => b[1] - a[1]);
    if (sorted.length > 0) {
      discoveredSelector = sorted[0][0];
      log('  DISCOVERED example selector: "' + discoveredSelector + '" (' + sorted[0][1] + ' matches)');
    }

    // Store first intent's examples immediately
    firstItem.examples = exampleCandidates.map(c => c.txt).join(', ');
    log('  First intent examples: "' + firstItem.examples.substring(0, 100) + '..."');
  } else {
    logWarn('  Could not find example text after clicking. Check the DEBUG logs above.');
    logWarn('  The script will still try broad text matching for remaining intents.');
  }

  // ── Step 5: Click each remaining intent and scrape ─────────────────────
  log('Step 4/5: Clicking each intent to extract examples...');
  const totalTime = Math.ceil((intentList.length * (CLICK_DELAY + BETWEEN_CLICKS)) / 1000 / 60);
  log('  Estimated time: ~' + totalTime + ' minutes for ' + intentList.length + ' intents.');

  // Start from index 1 since we already did index 0
  for (let i = 1; i < intentList.length; i++) {
    const item = intentList[i];

    // Take snapshot before click
    const snapBefore = snapshotLeafTexts();

    // Click the intent
    const treeContent = item._nodeEl.closest('.p-treenode-content') || item._nodeEl;
    treeContent.scrollIntoView({ block: 'center', behavior: 'instant' });
    treeContent.click();
    await sleep(CLICK_DELAY);

    // Take snapshot after click
    const snapAfter = snapshotLeafTexts();

    // Find new text elements
    const newItems = [];
    for (const [path, txt] of snapAfter) {
      if (!snapBefore.has(path)) {
        newItems.push({ path, txt });
      }
    }

    // If we have a discovered selector, also try to match by selector
    if (discoveredSelector) {
      const els = document.querySelectorAll(discoveredSelector);
      if (els.length > 0) {
        const texts = Array.from(els)
          .map(e => e.textContent.trim())
          .filter(t => t.length >= 10 && !/^\d+(\.\d+)?%?$/.test(t));
        if (texts.length > 0) {
          item.examples = texts.join(', ');
        }
      }
    }

    // If selector didn't work, use the diff approach
    if (!item.examples) {
      const candidates = newItems.filter(ni => {
        const t = ni.txt;
        if (t.length < 15 || t.length > 500) return false;
        if (/^\d+(\.\d+)?%?$/.test(t)) return false;
        if (t === item.intent || t === item.category || t === item.topic) return false;
        return true;
      });
      if (candidates.length > 0) {
        item.examples = candidates.map(c => c.txt).join(', ');
      }
    }

    const exCount = item.examples ? item.examples.split(', ').length : 0;
    logProgress(i + 1, intentList.length, item.category + ' > ' + item.topic + ' > ' + item.intent + ' (' + exCount + ' examples)');
    await sleep(BETWEEN_CLICKS);
  }

  // Log the first intent result (already scraped in discovery)
  logProgress(1, intentList.length, firstItem.category + ' > ' + firstItem.topic + ' > ' + firstItem.intent + ' (' + (firstItem.examples ? firstItem.examples.split(', ').length : 0) + ' examples)');

  // Clean up DOM references
  const rows = intentList.map(function(item) {
    return {
      category: item.category,
      topic: item.topic,
      intent: item.intent,
      intentPercentage: item.intentPercentage,
      volume: item.volume,
      examples: item.examples,
      active: item.active
    };
  });

  // ── Step 6: Download ───────────────────────────────────────────────────
  log('Step 5/5: Generating Excel file...');
  downloadExcel(rows);

  const categories = [];
  const seen = {};
  for (const r of rows) { if (!seen[r.category]) { seen[r.category] = 1; categories.push(r.category); } }
  const withExamples = rows.filter(function(r) { return r.examples; }).length;

  log('');
  log('=== COMPLETE ===');
  log('  Categories:      ' + categories.length);
  log('  Total intents:   ' + rows.length);
  log('  With examples:   ' + withExamples);
  log('  File downloaded: CXOne_Intents_Output.xlsx');

  console.table(rows.slice(0, 10));
  if (rows.length > 10) {
    log('  ... and ' + (rows.length - 10) + ' more rows (see the downloaded file).');
  }

  // ────────────────────────────────────────────────────────────────────────
  //  XLSX Generator + CSV + ZIP (zero dependencies)
  // ────────────────────────────────────────────────────────────────────────

  function downloadExcel(data) {
    var headers = ['Category','Topic','Intent','Intent Percentage','Volume','Examples','Active'];
    var keys = ['category','topic','intent','intentPercentage','volume','examples','active'];

    function esc(s) {
      if (s == null) return '';
      return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    }
    function col(i) {
      var s = ''; i++;
      while (i > 0) { i--; s = String.fromCharCode(65 + (i % 26)) + s; i = Math.floor(i / 26); }
      return s;
    }

    var sr = '<row r="1">';
    for (var c = 0; c < headers.length; c++) {
      sr += '<c r="' + col(c) + '1" t="inlineStr" s="1"><is><t>' + esc(headers[c]) + '</t></is></c>';
    }
    sr += '</row>';
    for (var r = 0; r < data.length; r++) {
      var rn = r + 2;
      sr += '<row r="' + rn + '">';
      for (var c2 = 0; c2 < keys.length; c2++) {
        sr += '<c r="' + col(c2) + rn + '" t="inlineStr"><is><t>' + esc(data[r][keys[c2]] || '') + '</t></is></c>';
      }
      sr += '</row>';
    }

    var lc = col(headers.length - 1);
    var lr = data.length + 1;

    var sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols><col min="1" max="1" width="25" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/><col min="3" max="3" width="45" customWidth="1"/><col min="4" max="4" width="18" customWidth="1"/><col min="5" max="5" width="12" customWidth="1"/><col min="6" max="6" width="60" customWidth="1"/><col min="7" max="7" width="10" customWidth="1"/></cols>' +
      '<sheetData>' + sr + '</sheetData>' +
      '<autoFilter ref="A1:' + lc + lr + '"/></worksheet>';

    var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill></fills>' +
      '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs>' +
      '<cellXfs count="2"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs></styleSheet>';

    var wbXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Intents" sheetId="1" r:id="rId1"/></sheets></workbook>';

    var wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';

    var rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';

    var ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';

    var blob = buildZip([
      {name:'[Content_Types].xml', content:ct},
      {name:'_rels/.rels', content:rels},
      {name:'xl/workbook.xml', content:wbXml},
      {name:'xl/_rels/workbook.xml.rels', content:wbRels},
      {name:'xl/styles.xml', content:stylesXml},
      {name:'xl/worksheets/sheet1.xml', content:sheetXml}
    ]);

    dl(blob, 'CXOne_Intents_Output.xlsx');
    log('  Excel file download triggered.');

    // CSV backup
    var csv = [headers.join(',')];
    for (var ri = 0; ri < data.length; ri++) {
      csv.push(keys.map(function(k) { return '"' + (data[ri][k]||'').replace(/"/g,'""') + '"'; }).join(','));
    }
    dl(new Blob([csv.join('\n')], {type:'text/csv'}), 'CXOne_Intents_Output.csv');
    log('  CSV backup also downloaded.');
  }

  function dl(blob, name) {
    var u = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = u; a.download = name;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(u);
  }

  function buildZip(files) {
    var enc = new TextEncoder();
    var parts = [], cd = [], off = 0;
    for (var i = 0; i < files.length; i++) {
      var nb = enc.encode(files[i].name);
      var cb = enc.encode(files[i].content);
      var cr = crc32(cb), sz = cb.length;

      var lh = new Uint8Array(30 + nb.length);
      var lv = new DataView(lh.buffer);
      lv.setUint32(0,0x04034b50,true); lv.setUint16(4,20,true);
      lv.setUint16(8,0,true); lv.setUint32(14,cr,true);
      lv.setUint32(18,sz,true); lv.setUint32(22,sz,true);
      lv.setUint16(26,nb.length,true); lh.set(nb,30);
      parts.push(lh, cb);

      var ce = new Uint8Array(46 + nb.length);
      var cv = new DataView(ce.buffer);
      cv.setUint32(0,0x02014b50,true); cv.setUint16(4,20,true); cv.setUint16(6,20,true);
      cv.setUint32(16,cr,true); cv.setUint32(20,sz,true); cv.setUint32(24,sz,true);
      cv.setUint16(28,nb.length,true); cv.setUint32(42,off,true);
      ce.set(nb,46); cd.push(ce);
      off += lh.length + cb.length;
    }
    var cdOff = off, cdSz = 0;
    for (var j = 0; j < cd.length; j++) { parts.push(cd[j]); cdSz += cd[j].length; }
    var eo = new Uint8Array(22);
    var ev = new DataView(eo.buffer);
    ev.setUint32(0,0x06054b50,true);
    ev.setUint16(8,files.length,true); ev.setUint16(10,files.length,true);
    ev.setUint32(12,cdSz,true); ev.setUint32(16,cdOff,true);
    parts.push(eo);
    return new Blob(parts,{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  }

  function crc32(b) {
    var c = 0xffffffff;
    for (var i = 0; i < b.length; i++) {
      c ^= b[i];
      for (var j = 0; j < 8; j++) c = (c >>> 1) ^ (c & 1 ? 0xedb88320 : 0);
    }
    return (c ^ 0xffffffff) >>> 0;
  }
})();
