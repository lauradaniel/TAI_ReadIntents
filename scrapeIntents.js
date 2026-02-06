const fs = require('fs');
const cheerio = require('cheerio');
const ExcelJS = require('exceljs');
const path = require('path');

const HTML_FILE = path.join(__dirname, 'CXOne_ReadIntents.txt');
const OUTPUT_FILE = path.join(__dirname, 'CXOne_Intents_Output.xlsx');

function scrapeIntents() {
  console.log('Reading HTML file...');
  const html = fs.readFileSync(HTML_FILE, 'utf-8');
  const $ = cheerio.load(html);

  // The HTML has a kanban-view with multiple kanban-tree elements (one per category).
  // Each tree has p-treenode elements with node-level-1 (Category), node-level-2 (Topic),
  // node-level-3 (Intent) classes, each containing a name and percentage.

  const rows = [];
  let currentCategory = '';
  let currentCategoryPct = '';
  let currentTopic = '';
  let currentTopicPct = '';

  // Select all kanban tree nodes in document order
  const nodes = $('.kanban-tree-node');
  console.log(`Found ${nodes.length} tree nodes`);

  nodes.each((_, el) => {
    const $el = $(el);
    const classList = $el.attr('class') || '';

    const name = $el.find('.kanban-tree-node-name').first().text().trim();
    const pctText = $el.find('.kanban-tree-node-statistics .percentage').first().text().trim();
    const percentage = pctText || '0%';

    if (classList.includes('node-level-1')) {
      currentCategory = name;
      currentCategoryPct = percentage;
      currentTopic = '';
      currentTopicPct = '';
    } else if (classList.includes('node-level-2')) {
      currentTopic = name;
      currentTopicPct = percentage;
    } else if (classList.includes('node-level-3')) {
      rows.push({
        category: currentCategory,
        topic: currentTopic,
        intent: name,
        intentPercentage: percentage,
        volume: '',       // Not available in static HTML (loaded dynamically)
        examples: '',     // Not available in static HTML (loaded dynamically)
        active: '',       // Not available in static HTML (loaded dynamically)
      });
    }
  });

  // Capture Topics that have no Level-3 children visible in the HTML.
  // The CXOne app loads intents lazily when a topic is expanded. If the topic
  // tree was collapsed when the HTML was saved, its Level-3 intents won't be
  // present. We still record the topic so nothing is silently lost.
  const topicsWithIntents = new Set();
  for (const row of rows) {
    topicsWithIntents.add(`${row.category}|${row.topic}`);
  }

  let cat = '', catPct = '';
  nodes.each((_, el) => {
    const $el = $(el);
    const classList = $el.attr('class') || '';
    const name = $el.find('.kanban-tree-node-name').first().text().trim();
    const pctText = $el.find('.kanban-tree-node-statistics .percentage').first().text().trim();

    if (classList.includes('node-level-1')) {
      cat = name;
      catPct = pctText;
    } else if (classList.includes('node-level-2')) {
      const key = `${cat}|${name}`;
      if (!topicsWithIntents.has(key)) {
        rows.push({
          category: cat,
          topic: name,
          intent: '(collapsed - expand topic in CXOne to capture intents)',
          intentPercentage: pctText || '0%',
          volume: '',
          examples: '',
          active: '',
        });
      }
    }
  });

  console.log(`Extracted ${rows.length} intent rows`);
  return rows;
}

async function writeExcel(rows) {
  console.log('Writing Excel file...');
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Intents');

  // Define columns
  sheet.columns = [
    { header: 'Category', key: 'category', width: 25 },
    { header: 'Topic', key: 'topic', width: 30 },
    { header: 'Intent', key: 'intent', width: 45 },
    { header: 'Intent Percentage', key: 'intentPercentage', width: 18 },
    { header: 'Volume', key: 'volume', width: 12 },
    { header: 'Examples', key: 'examples', width: 60 },
    { header: 'Active', key: 'active', width: 10 },
  ];

  // Style header row
  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4472C4' },
  };
  headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

  // Add data rows
  for (const row of rows) {
    sheet.addRow(row);
  }

  // Auto-filter
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: rows.length + 1, column: 7 },
  };

  // Freeze header row
  sheet.views = [{ state: 'frozen', ySplit: 1 }];

  await workbook.xlsx.writeFile(OUTPUT_FILE);
  console.log(`Excel file written to: ${OUTPUT_FILE}`);
}

async function main() {
  try {
    const rows = scrapeIntents();

    if (rows.length === 0) {
      console.error('No data extracted. Check the HTML file structure.');
      process.exit(1);
    }

    // Print summary
    const categories = [...new Set(rows.map(r => r.category))];
    console.log(`\nSummary:`);
    console.log(`  Categories: ${categories.length}`);
    console.log(`  Total rows: ${rows.length}`);
    console.log(`\nCategories found:`);
    for (const cat of categories) {
      const catRows = rows.filter(r => r.category === cat);
      console.log(`  ${cat}: ${catRows.length} intents`);
    }

    // Print first few rows as preview
    console.log('\nPreview (first 5 rows):');
    console.log('-'.repeat(120));
    for (const row of rows.slice(0, 5)) {
      console.log(`  ${row.category} > ${row.topic} > ${row.intent} (${row.intentPercentage})`);
    }
    console.log('-'.repeat(120));

    console.log('\nNote:');
    console.log('  - Volume, Examples, and Active columns are empty because this data');
    console.log('    is loaded dynamically in CXOne when clicking on an intent.');
    console.log('  - Topics marked "(collapsed)" had their tree collapsed when the HTML');
    console.log('    was saved. Expand all topics in CXOne before saving the HTML to');
    console.log('    capture all Level-3 intents.');

    await writeExcel(rows);
    console.log('\nDone!');
  } catch (err) {
    console.error('Error:', err.message);
    process.exit(1);
  }
}

main();
