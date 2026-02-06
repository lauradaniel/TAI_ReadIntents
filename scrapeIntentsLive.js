#!/usr/bin/env node

/**
 * CXOne Intent Scraper - Puppeteer (Live Browser) Version
 *
 * This script connects to a live CXOne session in your browser, expands all
 * Category / Topic / Intent trees, clicks each Intent to open its detail panel,
 * and extracts: Category, Topic, Intent, Intent Percentage, Volume, Examples, Active.
 *
 * Prerequisites:
 *   1. Install dependencies:   npm install
 *   2. Launch Chrome with remote debugging:
 *        Windows:  chrome.exe --remote-debugging-port=9222
 *        Mac:      /Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222
 *        Linux:    google-chrome --remote-debugging-port=9222
 *   3. In that Chrome window, log into CXOne and navigate to the Intent Builder
 *      (the page with the kanban tree view of Categories/Topics/Intents).
 *   4. Run this script:  node scrapeIntentsLive.js
 *
 * The script will:
 *   - Connect to your running Chrome instance
 *   - Find the CXOne tab
 *   - Expand every collapsed tree node (Category → Topic → Intent)
 *   - Click each Level-3 intent to load its detail panel
 *   - Extract percentage, volume, examples, and active status
 *   - Write everything to CXOne_Intents_Output.xlsx
 */

const puppeteer = require('puppeteer-core');
const ExcelJS = require('exceljs');
const path = require('path');

const OUTPUT_FILE = path.join(__dirname, 'CXOne_Intents_Output.xlsx');

// Configuration
const CONFIG = {
  // Chrome remote debugging URL
  browserURL: 'http://127.0.0.1:9222',
  // How long to wait after clicking an intent for the detail panel to load (ms)
  detailLoadDelay: 2000,
  // How long to wait after expanding a tree node (ms)
  expandDelay: 1000,
  // Max wait for selectors (ms)
  selectorTimeout: 10000,
};

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function safeClick(page, element) {
  try {
    await element.scrollIntoViewIfNeeded();
    await element.click();
    return true;
  } catch {
    return false;
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Step 1 – Expand all kanban tree nodes so every intent is visible
// ─────────────────────────────────────────────────────────────────────────────

async function expandAllTreeNodes(page) {
  console.log('\n[1/4] Expanding all collapsed tree nodes...');

  let expandedCount = 0;
  let pass = 0;

  // Keep expanding until no more collapsed togglers are found.
  // Each pass may reveal new children that are themselves collapsed.
  while (true) {
    pass++;
    // Find all collapsed p-tree toggler buttons (the chevron-right icons)
    // A collapsed node has a chevronrighticon inside its toggler
    const togglers = await page.$$('.kanban-view-panel .p-tree-toggler:has(chevronrighticon)');

    if (togglers.length === 0) {
      break;
    }

    console.log(`  Pass ${pass}: found ${togglers.length} collapsed node(s), expanding...`);

    for (const toggler of togglers) {
      await safeClick(page, toggler);
      await sleep(CONFIG.expandDelay);
      expandedCount++;
    }
  }

  console.log(`  Expanded ${expandedCount} tree node(s) across ${pass} pass(es).`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Step 2 – Collect all visible nodes with their hierarchy
// ─────────────────────────────────────────────────────────────────────────────

async function collectTreeNodes(page) {
  console.log('\n[2/4] Collecting all visible tree nodes...');

  const nodes = await page.evaluate(() => {
    const result = [];
    const els = document.querySelectorAll('.kanban-tree-node');

    for (const el of els) {
      const classList = el.className;
      let level = 0;
      if (classList.includes('node-level-1')) level = 1;
      else if (classList.includes('node-level-2')) level = 2;
      else if (classList.includes('node-level-3')) level = 3;
      if (level === 0) continue;

      const nameEl = el.querySelector('.kanban-tree-node-name');
      const pctEl = el.querySelector('.kanban-tree-node-statistics .percentage');

      result.push({
        level,
        name: nameEl ? nameEl.textContent.trim() : '',
        percentage: pctEl ? pctEl.textContent.trim() : '0%',
      });
    }
    return result;
  });

  // Build hierarchy: assign current category/topic to each intent
  const rows = [];
  let currentCategory = '';
  let currentTopic = '';

  for (const node of nodes) {
    if (node.level === 1) {
      currentCategory = node.name;
      currentTopic = '';
    } else if (node.level === 2) {
      currentTopic = node.name;
    } else if (node.level === 3) {
      rows.push({
        category: currentCategory,
        topic: currentTopic,
        intent: node.name,
        intentPercentage: node.percentage,
        volume: '',
        examples: '',
        active: '',
      });
    }
  }

  console.log(`  Found ${rows.length} Level-3 intents.`);
  return rows;
}

// ─────────────────────────────────────────────────────────────────────────────
// Step 3 – Click each intent and scrape the detail panel
// ─────────────────────────────────────────────────────────────────────────────

async function scrapeIntentDetails(page, rows) {
  console.log('\n[3/4] Clicking each intent to scrape detail panel...');
  console.log(`  Processing ${rows.length} intents (this may take a while)...\n`);

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const progress = `[${i + 1}/${rows.length}]`;

    // Find the Level-3 node element by matching its name text
    const clicked = await page.evaluate((intentName) => {
      const nodes = document.querySelectorAll('.kanban-tree-node.node-level-3');
      for (const node of nodes) {
        const nameEl = node.querySelector('.kanban-tree-node-name');
        if (nameEl && nameEl.textContent.trim() === intentName) {
          node.click();
          return true;
        }
      }
      return false;
    }, row.intent);

    if (!clicked) {
      console.log(`  ${progress} ${row.intent} — SKIPPED (node not found)`);
      continue;
    }

    // Wait for the detail panel to load
    await sleep(CONFIG.detailLoadDelay);

    // Scrape the detail panel data
    const detail = await page.evaluate(() => {
      const data = { volume: '', examples: '', active: '' };

      // Try multiple known selectors for the info panel
      // The CXOne app uses .info-body with .info-item children
      // each .info-item has a .sub-title and .item-value

      const infoItems = document.querySelectorAll('.info-body .info-item');
      for (const item of infoItems) {
        const titleEl = item.querySelector('.sub-title');
        const valueEl = item.querySelector('.item-value');
        if (!titleEl || !valueEl) continue;

        const title = titleEl.textContent.trim().toLowerCase();
        const value = valueEl.textContent.trim();

        if (title.includes('volume')) {
          data.volume = value;
        } else if (title.includes('example') || title.includes('sample') || title.includes('utterance') || title.includes('training')) {
          data.examples = value;
        } else if (title.includes('active') || title.includes('status')) {
          data.active = value;
        }
      }

      // Alternative: Try to find examples in a list/table format
      if (!data.examples) {
        // Check for sample sentences in ag-grid or list within the detail view
        const sampleRows = document.querySelectorAll(
          'item-info-panel .sample-list .sample-item, ' +
          'item-info-panel .example-row, ' +
          '.intent-examples .example-text, ' +
          '.info-body .examples-list li, ' +
          '.sentence-list .sentence-item'
        );
        if (sampleRows.length > 0) {
          data.examples = Array.from(sampleRows)
            .map(el => el.textContent.trim())
            .filter(Boolean)
            .join('\n');
        }
      }

      // Alternative: check for percentage & volume in the panel header area
      if (!data.volume) {
        // Sometimes volume is shown near the percentage in the detail header
        const headerStats = document.querySelectorAll('.panel-header .stat-value, .intent-stats .volume');
        for (const stat of headerStats) {
          const text = stat.textContent.trim();
          if (text && /\d/.test(text)) {
            data.volume = text;
            break;
          }
        }
      }

      // Check for active status via toggle/checkbox in the detail panel
      if (!data.active) {
        const toggle = document.querySelector(
          'item-info-panel .active-toggle input, ' +
          'item-info-panel p-checkbox input, ' +
          '.info-body .active-status'
        );
        if (toggle) {
          data.active = toggle.checked ? 'Yes' : 'No';
        }
      }

      return data;
    });

    row.volume = detail.volume;
    row.examples = detail.examples;
    row.active = detail.active;

    const examplePreview = detail.examples
      ? ` (${detail.examples.split('\n').length} examples)`
      : '';
    console.log(`  ${progress} ${row.category} > ${row.topic} > ${row.intent}${examplePreview}`);
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Step 4 – Write Excel
// ─────────────────────────────────────────────────────────────────────────────

async function writeExcel(rows) {
  console.log('\n[4/4] Writing Excel file...');

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Intents');

  sheet.columns = [
    { header: 'Category', key: 'category', width: 25 },
    { header: 'Topic', key: 'topic', width: 30 },
    { header: 'Intent', key: 'intent', width: 45 },
    { header: 'Intent Percentage', key: 'intentPercentage', width: 18 },
    { header: 'Volume', key: 'volume', width: 12 },
    { header: 'Examples', key: 'examples', width: 60 },
    { header: 'Active', key: 'active', width: 10 },
  ];

  // Style header
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

  // Add data
  for (const row of rows) {
    const excelRow = sheet.addRow(row);
    // Wrap text in Examples column
    excelRow.getCell('examples').alignment = { wrapText: true, vertical: 'top' };
  }

  // Auto-filter & freeze
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rows.length + 1, column: 7 },
  };
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  await workbook.xlsx.writeFile(OUTPUT_FILE);
  console.log(`  Excel file written to: ${OUTPUT_FILE}`);
}

// ─────────────────────────────────────────────────────────────────────────────
// Main
// ─────────────────────────────────────────────────────────────────────────────

async function main() {
  console.log('=== CXOne Intent Scraper (Live Browser) ===\n');
  console.log('Connecting to Chrome at', CONFIG.browserURL, '...');

  let browser;
  try {
    browser = await puppeteer.connect({ browserURL: CONFIG.browserURL });
  } catch (err) {
    console.error('\nFailed to connect to Chrome. Make sure Chrome is running with:');
    console.error('  chrome --remote-debugging-port=9222\n');
    console.error('Then log into CXOne and navigate to the Intent Builder page.\n');
    console.error('Error:', err.message);
    process.exit(1);
  }

  // Find the CXOne tab
  const pages = await browser.pages();
  let page = null;
  for (const p of pages) {
    const url = p.url();
    if (url.includes('cxone') || url.includes('nice') || url.includes('intent')) {
      page = p;
      break;
    }
  }

  if (!page) {
    // Fall back to the active tab
    page = pages[pages.length - 1];
    console.log('  Could not auto-detect CXOne tab, using the last active tab.');
    console.log('  Current URL:', page.url());
  } else {
    console.log('  Found CXOne tab:', page.url());
  }

  // Verify we see the kanban view
  const hasKanban = await page.$('.kanban-view-panel');
  if (!hasKanban) {
    console.error('\nCould not find the kanban view on this page.');
    console.error('Make sure you are on the CXOne Intent Builder page with the');
    console.error('Category / Topic / Intent tree visible.');
    process.exit(1);
  }

  try {
    // Step 1: Expand all trees
    await expandAllTreeNodes(page);

    // Step 2: Collect the tree hierarchy
    const rows = await collectTreeNodes(page);

    if (rows.length === 0) {
      console.error('No intents found. Make sure the kanban tree is visible.');
      process.exit(1);
    }

    // Step 3: Click each intent and scrape details
    await scrapeIntentDetails(page, rows);

    // Step 4: Write to Excel
    await writeExcel(rows);

    // Summary
    const withExamples = rows.filter((r) => r.examples).length;
    const categories = [...new Set(rows.map((r) => r.category))];
    console.log('\n=== Summary ===');
    console.log(`  Categories:    ${categories.length}`);
    console.log(`  Total intents: ${rows.length}`);
    console.log(`  With examples: ${withExamples}`);
    console.log(`  Output file:   ${OUTPUT_FILE}`);
    console.log('\nDone!');
  } catch (err) {
    console.error('Error during scraping:', err.message);
    process.exit(1);
  }
}

main();
