/* test.js — Sheet Shaper surgical-clone integration test
 *
 * Run with:  node test.js   (or: npm test)
 *
 * Builds a Reef-Desk-shaped master in-memory via ExcelJS, extracts the
 * SURGICAL CLONE + DATE HELPERS blocks from sheet-shaper.html, runs them
 * under vm.runInContext with @xmldom/xmldom-backed DOM, and asserts the
 * output zip is structurally clean — i.e. that Excel won't pop the
 * "We found a problem with some content" dialog on open.
 */

'use strict';

const fs = require('fs');
const os = require('os');
const path = require('path');
const vm = require('vm');
const assert = require('assert');
const { execSync } = require('child_process');

const JSZip = require('jszip');
const ExcelJS = require('exceljs');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

const HTML_PATH = path.join(__dirname, 'sheet-shaper.html');

// ---------------------------------------------------------------------------
// Step 1: extract production code from sheet-shaper.html (no duplication)
// ---------------------------------------------------------------------------

function extractBlock(html, beginRe, endRe) {
  const start = html.search(beginRe);
  const end = html.search(endRe);
  if (start < 0 || end < 0 || end <= start) {
    throw new Error(`could not extract block: begin=${start} end=${end}`);
  }
  const afterBegin = html.indexOf('\n', start);
  return html.slice(afterBegin + 1, end);
}

function loadGenerator() {
  const html = fs.readFileSync(HTML_PATH, 'utf8');
  const dateHelpers = extractBlock(
    html,
    /\/\* === DATE HELPERS BEGIN === \*\//,
    /\/\* === DATE HELPERS END === \*\//,
  );
  const surgical = extractBlock(
    html,
    /\/\* === SURGICAL CLONE BEGIN === \*\//,
    /\/\* === SURGICAL CLONE END === \*\//,
  );
  // Constants that live above the helper block in the HTML — mirror them
  // here so the sandbox can resolve MONTH_LONG / DAY_LONG / MONTH_SHORT.
  const constants = `
    const MONTH_LONG = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    const MONTH_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const DAY_LONG = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];
  `;
  const sandbox = {
    JSZip, DOMParser, XMLSerializer, console,
    // STATE is referenced by setupOutputDirectory; not used in our test path
    // but the function definition references it lexically.
    STATE: { mode: 'files', dirHandle: null },
    toast: () => {},
  };
  vm.createContext(sandbox);
  vm.runInContext(constants + '\n' + dateHelpers + '\n' + surgical + '\n', sandbox, { filename: 'sheet-shaper-extracted.js' });
  if (typeof sandbox.generateSurgicalWorkbook !== 'function') {
    throw new Error('generateSurgicalWorkbook not defined after extraction');
  }
  return sandbox;
}

// ---------------------------------------------------------------------------
// Step 2: synthesize a Reef-Desk-shaped master in-memory
// ---------------------------------------------------------------------------

async function buildSyntheticMaster() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'sheet-shaper-test';
  const ws = wb.addWorksheet('Template');

  // Title row: merged A1:H1 with a placeholder
  ws.mergeCells('A1:H1');
  ws.getCell('A1').value = 'Excursion — Date: {{date}}';

  // Date-derivation formula at A2 (the one ExcelJS mangles in the wild)
  ws.getCell('A2').value = {
    formula: 'DATEVALUE(MID(CELL("filename",A2),FIND("]",CELL("filename",A2))+1,255)&"/01/2024")',
  };

  // Header row at row 4
  ws.getCell('A4').value = 'Room';
  ws.getCell('B4').value = 'Guest';
  ws.getCell('C4').value = 'Attended';
  ws.getCell('D4').value = 'Pax';
  ws.getCell('E4').value = 'Adults';
  ws.getCell('F4').value = 'Children';
  ws.getCell('G4').value = 'Sub';
  ws.getCell('H4').value = 'Total';

  // Shared/filled-down formula across rows 5..24 (G = E + F)
  ws.fillFormula('G5:G24', 'E5+F5');

  // A second formula column to exercise the shared-formula path
  ws.fillFormula('H5:H24', 'G5*2');

  // Data validation dropdown
  ws.dataValidations.add('C5:C24', {
    type: 'list',
    allowBlank: true,
    formulae: ['"Yes,No"'],
  });

  // Conditional formatting: pax > 4 → bold
  ws.addConditionalFormatting({
    ref: 'D5:D24',
    rules: [{
      type: 'expression',
      formulae: ['D5>4'],
      style: { font: { bold: true } },
    }],
  });

  // Extra merged range outside the title
  ws.mergeCells('A26:C27');
  ws.getCell('A26').value = 'Notes: {{date}}';

  return await wb.xlsx.writeBuffer();
}

// ---------------------------------------------------------------------------
// Step 3: assertion helpers
// ---------------------------------------------------------------------------

function colLettersToNum(s) {
  let n = 0;
  for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

function parseCellRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) throw new Error(`bad cell ref: ${ref}`);
  return { col: colLettersToNum(m[1]), row: parseInt(m[2], 10) };
}

async function runXmllint(zip) {
  // xmllint --noout every .xml / .rels part. Throws if any fail to parse.
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-shaper-xmllint-'));
  try {
    const xmlNames = Object.keys(zip.files).filter(n => /\.(xml|rels)$/.test(n) && !zip.files[n].dir);
    for (const name of xmlNames) {
      const content = await zip.file(name).async('nodebuffer');
      const tmpPath = path.join(tmpDir, name.replace(/[\/]/g, '_'));
      fs.writeFileSync(tmpPath, content);
      try {
        execSync(`xmllint --noout ${JSON.stringify(tmpPath)}`, { stdio: 'pipe' });
      } catch (e) {
        const stderr = (e.stderr || Buffer.alloc(0)).toString();
        throw new Error(`xmllint failed on ${name}:\n${stderr}`);
      }
    }
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

function commandExists(cmd) {
  try {
    execSync(`command -v ${cmd}`, { stdio: 'pipe' });
    return true;
  } catch { return false; }
}

async function buildSofficeBaseline() {
  // Sanity-check: ensure soffice can load a known-good xlsx in this env.
  // If it can't, LibreOffice itself is misconfigured (java/profile/etc) and
  // a failure on our output isn't diagnostic.
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('A').getCell('A1').value = 'baseline';
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function runLibreOfficeRoundtrip(outBuf) {
  if (!commandExists('soffice')) {
    console.log('  · soffice not on PATH — skipping LibreOffice round-trip');
    return;
  }
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sheet-shaper-lo-'));
  const baselineDir = path.join(tmpDir, 'baseline');
  const outDir = path.join(tmpDir, 'out');
  fs.mkdirSync(baselineDir);
  fs.mkdirSync(outDir);
  const baselineIn = path.join(tmpDir, 'baseline.xlsx');
  const inPath = path.join(tmpDir, 'in.xlsx');
  fs.writeFileSync(baselineIn, await buildSofficeBaseline());
  fs.writeFileSync(inPath, Buffer.from(outBuf));

  const sofficeOk = (input, outDir) => {
    try {
      execSync(
        `soffice --headless --convert-to xlsx --outdir ${JSON.stringify(outDir)} ${JSON.stringify(input)}`,
        { stdio: 'pipe', timeout: 90000, env: { ...process.env, HOME: tmpDir } },
      );
    } catch (e) {
      const stderr = (e.stderr || Buffer.alloc(0)).toString();
      throw new Error(`soffice errored: ${e.message}\nstderr: ${stderr}`);
    }
    return fs.readdirSync(outDir);
  };

  try {
    const baselineOut = sofficeOk(baselineIn, baselineDir);
    if (baselineOut.length === 0) {
      console.log('  · soffice cannot convert a vanilla xlsx in this env — skipping (env issue, not our bug)');
      return;
    }
    const produced = sofficeOk(inPath, outDir);
    assert.ok(produced.length > 0, 'LibreOffice produced no output file for our xlsx (likely corruption)');
    const stat = fs.statSync(path.join(outDir, produced[0]));
    assert.ok(stat.size > 0, 'LibreOffice output is empty');
    console.log(`  · LibreOffice round-trip OK (${produced[0]}, ${stat.size} bytes)`);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

// ---------------------------------------------------------------------------
// Step 4: the test
// ---------------------------------------------------------------------------

async function main() {
  console.log('1. Extracting surgical-clone block from sheet-shaper.html');
  const ctx = loadGenerator();

  console.log('2. Building synthetic Reef-Desk master via ExcelJS');
  const masterArrayBuffer = await buildSyntheticMaster();
  // ExcelJS returns a Node Buffer-like; coerce to ArrayBuffer for parity
  // with what the browser's File.arrayBuffer() returns.
  const masterBuf = masterArrayBuffer instanceof ArrayBuffer
    ? masterArrayBuffer
    : Uint8Array.from(masterArrayBuffer).buffer;

  // Also crack the master open to confirm it has the bits we synthesized.
  const masterZip = await JSZip.loadAsync(masterBuf);
  const masterSheetNames = Object.keys(masterZip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  assert.strictEqual(masterSheetNames.length, 1, `master should have 1 sheet, has ${masterSheetNames.length}`);
  const masterSheetXml = await masterZip.file(masterSheetNames[0]).async('string');
  // ExcelJS XML-escapes embedded quotes as &quot; — accept either form.
  const cellFormulaRe = /CELL\((?:"|&quot;)filename(?:"|&quot;)/;
  assert.ok(cellFormulaRe.test(masterSheetXml), 'master sheet missing CELL("filename") formula');
  assert.ok(/<mergeCells\b/.test(masterSheetXml), 'master sheet missing mergeCells');
  assert.ok(/<dataValidations\b/.test(masterSheetXml), 'master sheet missing dataValidations');
  assert.ok(/<conditionalFormatting\b/.test(masterSheetXml), 'master sheet missing conditionalFormatting');

  console.log('3. Running surgical generator for 5 May 2026 dates');
  const dates = [1, 2, 3, 4, 5].map(d => new Date(2026, 4, d));
  const master = { rawBuf: masterBuf, excursionName: 'Test Excursion' };
  const outBuf = await ctx.generateSurgicalWorkbook(master, dates, 'D');
  assert.ok(outBuf, 'generator returned nothing');
  assert.ok(outBuf.byteLength > 0, 'generator returned empty buffer');

  console.log('4. Inspecting output zip');
  const outZip = await JSZip.loadAsync(outBuf);

  // A. Exactly 5 sheetN.xml files
  const sheetFiles = Object.keys(outZip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
  assert.deepStrictEqual(
    sheetFiles,
    ['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/worksheets/sheet3.xml','xl/worksheets/sheet4.xml','xl/worksheets/sheet5.xml'],
    `expected 5 sheets, got: ${sheetFiles.join(', ')}`,
  );

  // B. Each contains CELL("filename" verbatim, plus dataValidations, plus
  //    conditionalFormatting, plus mergeCells. And the mergeCells block is
  //    byte-identical across all 5 sheets.
  const mergesByFile = {};
  const dvByFile = {};
  const cfByFile = {};
  for (let i = 1; i <= 5; i++) {
    const xml = await outZip.file(`xl/worksheets/sheet${i}.xml`).async('string');
    assert.ok(cellFormulaRe.test(xml), `sheet${i} missing CELL("filename") formula`);
    const m = /<mergeCells\b[\s\S]*?<\/mergeCells>/.exec(xml);
    assert.ok(m, `sheet${i} missing <mergeCells>`);
    mergesByFile[i] = m[0];
    const dv = /<dataValidations\b[\s\S]*?<\/dataValidations>/.exec(xml);
    assert.ok(dv, `sheet${i} missing <dataValidations>`);
    dvByFile[i] = dv[0];
    const cf = /<conditionalFormatting\b[\s\S]*?<\/conditionalFormatting>/.exec(xml);
    assert.ok(cf, `sheet${i} missing <conditionalFormatting>`);
    cfByFile[i] = cf[0];
  }
  for (let i = 2; i <= 5; i++) {
    assert.strictEqual(mergesByFile[i], mergesByFile[1], `mergeCells differs between sheet1 and sheet${i}`);
    assert.strictEqual(dvByFile[i], dvByFile[1], `dataValidations differs between sheet1 and sheet${i}`);
    assert.strictEqual(cfByFile[i], cfByFile[1], `conditionalFormatting differs between sheet1 and sheet${i}`);
  }

  // C. workbook.xml lists 5 sheets named "1".."5"
  const wbXml = await outZip.file('xl/workbook.xml').async('string');
  const sheetsBlock = /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/.exec(wbXml);
  assert.ok(sheetsBlock, 'workbook.xml missing <sheets> block');
  const sheetNameMatches = [...sheetsBlock[1].matchAll(/<sheet\s+name="([^"]+)"/g)].map(m => m[1]);
  assert.deepStrictEqual(sheetNameMatches, ['1','2','3','4','5'], `sheet names: ${sheetNameMatches.join(',')}`);

  // D. [Content_Types].xml has Overrides for all 5 sheets AND original ones
  const ctXml = await outZip.file('[Content_Types].xml').async('string');
  for (let i = 1; i <= 5; i++) {
    assert.ok(
      ctXml.includes(`PartName="/xl/worksheets/sheet${i}.xml"`),
      `content types missing sheet${i} override`,
    );
  }
  assert.ok(/PartName="\/xl\/styles\.xml"/.test(ctXml), 'content types missing styles override');
  assert.ok(/PartName="\/xl\/sharedStrings\.xml"/.test(ctXml), 'content types missing sharedStrings override');
  // theme is optional (ExcelJS doesn't always emit one) — assert if present in master
  const masterCtXml = await masterZip.file('[Content_Types].xml').async('string');
  if (/PartName="\/xl\/theme\/theme1\.xml"/.test(masterCtXml)) {
    assert.ok(/PartName="\/xl\/theme\/theme1\.xml"/.test(ctXml), 'content types lost theme override');
  }

  // E. calcChain gone everywhere
  assert.ok(!outZip.file('xl/calcChain.xml'), 'xl/calcChain.xml should be absent');
  assert.ok(!/calcChain/.test(ctXml), '[Content_Types].xml still references calcChain');
  const wbRelsXml = await outZip.file('xl/_rels/workbook.xml.rels').async('string');
  assert.ok(!/calcChain/.test(wbRelsXml), 'workbook.xml.rels still references calcChain');

  // F. workbook rels: one relationship per sheet, original styles/theme/sharedStrings preserved
  for (let i = 1; i <= 5; i++) {
    assert.ok(
      wbRelsXml.includes(`Target="worksheets/sheet${i}.xml"`),
      `workbook rels missing sheet${i}`,
    );
  }
  assert.ok(/styles\.xml/.test(wbRelsXml), 'workbook rels lost styles');
  assert.ok(/sharedStrings\.xml/.test(wbRelsXml), 'workbook rels lost sharedStrings');

  // G. No non-anchor cell inside any merge range appears as its own <c>
  for (let i = 1; i <= 5; i++) {
    const xml = await outZip.file(`xl/worksheets/sheet${i}.xml`).async('string');
    const mergeBlock = /<mergeCells\b[\s\S]*?<\/mergeCells>/.exec(xml);
    if (!mergeBlock) continue;
    const mergeRefs = [...mergeBlock[0].matchAll(/<mergeCell\s+ref="([^"]+)"/g)].map(m => m[1]);
    const nonAnchors = new Set();
    for (const ref of mergeRefs) {
      const [start, end] = ref.split(':');
      const s = parseCellRef(start);
      const e = end ? parseCellRef(end) : s;
      for (let r = s.row; r <= e.row; r++) {
        for (let c = s.col; c <= e.col; c++) {
          if (r === s.row && c === s.col) continue;
          nonAnchors.add(`${r},${c}`);
        }
      }
    }
    // Only value-carrying cells inside a merge range trigger Excel's repair.
    // Empty <c r="B1"/> stubs are harmless and standard ExcelJS output.
    const cellMatches = [...xml.matchAll(/<c\s+r="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>|<c\s+r="[A-Z]+\d+"[^>]*\/>/g)];
    const valueCells = [];
    const cellWithContentRe = /<c\s+r="([A-Z]+\d+)"[^>]*>(?:\s*<(?:v|f|is)\b)/g;
    let mc;
    while ((mc = cellWithContentRe.exec(xml)) !== null) valueCells.push(mc[1]);
    for (const cellRef of valueCells) {
      const p = parseCellRef(cellRef);
      assert.ok(
        !nonAnchors.has(`${p.row},${p.col}`),
        `sheet${i}: non-anchor merge cell ${cellRef} has content (would trigger Excel repair)`,
      );
    }
  }

  // H. Placeholder substitution happened — A1 cell should contain a real date
  for (let i = 1; i <= 5; i++) {
    const xml = await outZip.file(`xl/worksheets/sheet${i}.xml`).async('string');
    assert.ok(
      !/\{\{\s*date\s*\}\}/.test(xml),
      `sheet${i} still contains a literal {{date}} placeholder`,
    );
    // A1 is the title; check it has the substituted long-form date
    assert.ok(
      xml.includes(`${i} May 2026`),
      `sheet${i} title missing substituted date "${i} May 2026"`,
    );
  }

  console.log('5. xmllint --noout on every XML part');
  await runXmllint(outZip);

  console.log('6. LibreOffice round-trip (if available)');
  await runLibreOfficeRoundtrip(outBuf);

  console.log('\nAll assertions passed.');
}

main().catch(e => {
  console.error('\nTEST FAILED:', e && e.stack ? e.stack : e);
  process.exit(1);
});
