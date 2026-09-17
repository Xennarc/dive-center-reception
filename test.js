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
const ENGINE_PATH = path.join(__dirname, 'Booking engine v5.html');

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
  // Expose extracted source for static-analysis assertions (e.g. catching
  // forbidden createElement( usage in namespaced contexts).
  sandbox.__extractedSurgical = surgical;
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

  // I. Namespace correctness — the regression that broke PR #29.
  //
  //    The browser's W3C-correct XMLSerializer emits xmlns="" on any element
  //    that's in NO namespace under a default-namespaced parent. Excel
  //    rejects the resulting file ("file format or file extension is not
  //    valid") because, e.g., <Override xmlns="" .../> doesn't bind a
  //    content type for our worksheets. @xmldom/xmldom silently drops the
  //    namespace mismatch on serialization, hiding the bug in Node, so we
  //    can't rely on the output alone — we also static-scan the production
  //    code we extracted to ensure no createElement( calls remain in the
  //    namespaced-XML rewrite paths.
  console.log('  · namespace correctness:');

  // I.1 Static scan of the extracted surgical-clone source.
  const extractedSrc = ctx.__extractedSurgical || '';
  assert.ok(extractedSrc.length, 'extracted surgical source missing — bug in test harness');
  const stripped = extractedSrc.replace(/createElementNS\s*\(/g, ''); // mask legitimate calls
  assert.ok(
    !/\.createElement\s*\(/.test(stripped),
    'createElement( found in surgical-clone block — use createElementNS so the new node inherits the parent\'s default namespace',
  );

  // I.2 Output string check — no xmlns="" anywhere in the rewritten parts.
  const checkXmlnsEmpty = async (partName) => {
    const xml = await outZip.file(partName).async('string');
    assert.ok(
      !xml.includes('xmlns=""'),
      `${partName} contains xmlns="" — browser would emit this for any element appended without inheriting the parent's default namespace`,
    );
  };
  await checkXmlnsEmpty('[Content_Types].xml');
  await checkXmlnsEmpty('xl/_rels/workbook.xml.rels');
  for (let i = 1; i <= 5; i++) await checkXmlnsEmpty(`xl/worksheets/sheet${i}.xml`);

  // I.3 Re-parse and verify every Override / Relationship / is / t we
  //     emitted is in the expected namespace.
  const NS_CT   = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const NS_REL  = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const NS_SS   = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const ctDoc = new DOMParser().parseFromString(
    await outZip.file('[Content_Types].xml').async('string'),
    'application/xml',
  );
  const ctOverrides = ctDoc.getElementsByTagName('Override');
  for (let i = 0; i < ctOverrides.length; i++) {
    assert.strictEqual(
      ctOverrides[i].namespaceURI, NS_CT,
      `Override[${i}] PartName=${ctOverrides[i].getAttribute('PartName')} has wrong namespaceURI (${ctOverrides[i].namespaceURI})`,
    );
  }
  const relDoc = new DOMParser().parseFromString(
    await outZip.file('xl/_rels/workbook.xml.rels').async('string'),
    'application/xml',
  );
  const relRels = relDoc.getElementsByTagName('Relationship');
  for (let i = 0; i < relRels.length; i++) {
    assert.strictEqual(
      relRels[i].namespaceURI, NS_REL,
      `Relationship[${i}] Id=${relRels[i].getAttribute('Id')} has wrong namespaceURI (${relRels[i].namespaceURI})`,
    );
  }
  for (let i = 1; i <= 5; i++) {
    const sDoc = new DOMParser().parseFromString(
      await outZip.file(`xl/worksheets/sheet${i}.xml`).async('string'),
      'application/xml',
    );
    const inlineCells = [];
    const allCells = sDoc.getElementsByTagName('c');
    for (let j = 0; j < allCells.length; j++) {
      if (allCells[j].getAttribute('t') === 'inlineStr') inlineCells.push(allCells[j]);
    }
    assert.ok(inlineCells.length > 0, `sheet${i} should have at least one inlineStr cell (the substituted title)`);
    for (const c of inlineCells) {
      const is = c.getElementsByTagName('is')[0];
      assert.ok(is, `sheet${i} inlineStr cell ${c.getAttribute('r')} missing <is>`);
      assert.strictEqual(is.namespaceURI, NS_SS, `sheet${i} <is> in cell ${c.getAttribute('r')} has wrong namespaceURI (${is.namespaceURI})`);
      const tNode = is.getElementsByTagName('t')[0];
      assert.ok(tNode, `sheet${i} <is> missing <t>`);
      assert.strictEqual(tNode.namespaceURI, NS_SS, `sheet${i} <t> in cell ${c.getAttribute('r')} has wrong namespaceURI (${tNode.namespaceURI})`);
    }
  }

  console.log('5. xmllint --noout on every XML part');
  await runXmllint(outZip);

  console.log('6. LibreOffice round-trip (if available)');
  await runLibreOfficeRoundtrip(outBuf);

  console.log('\nAll assertions passed.');
}

// ---------------------------------------------------------------------------
// Step 5: multi-boat helpers extracted from Booking engine v5.html
// ---------------------------------------------------------------------------

function loadBoatHelpers() {
  const html = fs.readFileSync(ENGINE_PATH, 'utf8');
  const boatInfer = extractBlock(html, /\/\* === BOAT_INFER_BEGIN === \*\//, /\/\* === BOAT_INFER_END === \*\//);
  const ensureCap = extractBlock(html, /\/\* === ENSURE_CAPACITY_BEGIN === \*\//, /\/\* === ENSURE_CAPACITY_END === \*\//);
  const rebuild   = extractBlock(html, /\/\* === REBUILD_GROUPINGS_BEGIN === \*\//, /\/\* === REBUILD_GROUPINGS_END === \*\//);
  const bulk      = extractBlock(html, /\/\* === BULK_ASSIGN_BEGIN === \*\//, /\/\* === BULK_ASSIGN_END === \*\//);

  // Stubs for v5 helpers the extracted code calls. Kept minimal — these
  // are tiny utilities; their browser-side definitions are not under test.
  const stubs = `
    function cellRaw(cell) {
      if (!cell) return null;
      const v = cell.value;
      if (v == null) return null;
      if (v instanceof Date) return v;
      if (typeof v === 'object' && 'result' in v) return v.result;
      if (typeof v === 'object' && 'text' in v) return v.text;
      if (typeof v === 'object' && 'richText' in v) return v.richText.map(r => r.text).join('');
      return v;
    }
    function cellStr(cell) {
      const v = cellRaw(cell);
      if (v == null) return '';
      return String(v).trim();
    }
    function parseDateValue(v) {
      if (!v) return null;
      if (v instanceof Date) {
        const yy = v.getFullYear(), mm = String(v.getMonth()+1).padStart(2,'0'), dd = String(v.getDate()).padStart(2,'0');
        return yy + '-' + mm + '-' + dd;
      }
      if (typeof v === 'string') {
        const m = v.match(/^(\\d{4})-(\\d{2})-(\\d{2})/);
        if (m) return m[1] + '-' + m[2] + '-' + m[3];
      }
      return null;
    }
    function toDeparture(s) {
      if (!s) return null;
      const iso = parseDateValue(s);
      if (iso) { const [y,m,d] = iso.split('-').map(Number); return new Date(y, m-1, d); }
      return s;
    }
    function readBooking(ws, sheetName, fileKey, layout, rowIdx) {
      // Minimal reader for ensureCapacity / rebuildBoatGroupings tests.
      const row = ws.getRow(rowIdx);
      const cols = layout.layout;
      const at = (c) => c == null ? null : row.getCell(c);
      const room = cellStr(at(cols.room));
      const guest = cellStr(at(cols.guest));
      const departureRaw = cellRaw(at(cols.departure));
      const departure = parseDateValue(departureRaw);
      const adults = (() => { const v = cellRaw(at(cols.adults)); return v == null || v === '' ? null : Number(v); })();
      const child = (() => { const v = cellRaw(at(cols.child)); return v == null || v === '' ? null : Number(v); })();
      const liability = cellStr(at(cols.liability));
      const billNo = cellStr(at(cols.billNo));
      const remarks = cellStr(at(cols.remarks));
      const staff = cellStr(at(cols.staff));
      const isEmpty = !room && !guest && adults == null && child == null && !liability && !billNo && !remarks && !staff && !departureRaw;
      const depStr = typeof departureRaw === 'string' ? departureRaw : '';
      const isSeparator = !room && /^\\s*BOAT\\s*\\d+\\s*$/i.test(depStr) && !!guest;
      return { fileKey, sheetName, rowIdx, room, guest, departure, departureRaw, adults, child, liability, billNo, remarks, staff, isEmpty, isSeparator };
    }
    // Used by rebuildBoatGroupings → getSheetMeta(file, name). Returns the
    // map entry the test seeds onto fileEntry.sheetMeta.
    function getSheetMeta(fileEntry, sheetName) {
      const stored = (fileEntry && fileEntry.sheetMeta && fileEntry.sheetMeta.get(sheetName)) || {};
      return Object.assign({ boats: [], assignments: {} }, stored);
    }
    function setSheetMeta(fileEntry, sheetName, partial) {
      if (!fileEntry.sheetMeta) fileEntry.sheetMeta = new Map();
      const existing = fileEntry.sheetMeta.get(sheetName) || {};
      fileEntry.sheetMeta.set(sheetName, Object.assign({}, existing, partial));
      fileEntry.dirty = true;
      if (!fileEntry.pendingOps) fileEntry.pendingOps = [];
      fileEntry.pendingOps.push({ kind: 'setSheetMeta', sheetName, fields: partial });
    }
    function _setBoatsAndAssignments(fileEntry, sheetName, boats, assignments) {
      setSheetMeta(fileEntry, sheetName, { boats, assignments });
    }
    function isMultiBoatFile(fileEntry) { return fileEntry && fileEntry.__multiBoat === true; }
    function rebuildBookings() {}
    function updateStatus() {}
    function recordOp() {}
    function excursionKeyFor() { return null; }
    function getExcursionConfig() { return { boatPool: [] }; }
    function saveExcursionConfig() {}
    function toast(msg, kind) { (collectedToasts.push({ msg, kind })); }
    const collectedToasts = [];
    const STATE = { warnedSumFormula: new Set(), files: [], bookings: [] };
  `;

  const sandbox = { console };
  vm.createContext(sandbox);
  vm.runInContext(stubs + '\n' + boatInfer + '\n' + ensureCap + '\n' + rebuild + '\n' + bulk + '\n', sandbox, { filename: 'engine-extracted.js' });
  if (typeof sandbox.inferBoatsFromSheet !== 'function') throw new Error('inferBoatsFromSheet missing');
  if (typeof sandbox.ensureCapacity !== 'function') throw new Error('ensureCapacity missing');
  if (typeof sandbox.rebuildBoatGroupings !== 'function') throw new Error('rebuildBoatGroupings missing');
  if (typeof sandbox.bulkAssignToBoat !== 'function') throw new Error('bulkAssignToBoat missing');
  return sandbox;
}

async function testBoatInference(ctx) {
  // Build a Sunset-Dolphin-shaped sheet: headerRow 6, firstData 8, total at 18.
  // Cols: room=4, guest=5, departure=6, adults=7, child=8.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('01');
  ws.getRow(6).getCell(4).value = 'Room';
  ws.getRow(6).getCell(5).value = 'Guest';
  ws.getRow(6).getCell(6).value = 'Departure';
  // BOAT 1 separator at row 7 (but firstData is 8, separator lives inside data range)
  // Per the real Sunset Dolphin layout: separator at row 7, data starts at row 8.
  ws.getRow(7).getCell(5).value = 'SHADHA';
  ws.getRow(7).getCell(6).value = 'BOAT 1';
  ws.getRow(8).getCell(4).value = '101'; ws.getRow(8).getCell(5).value = 'Alice'; ws.getRow(8).getCell(6).value = '2026-05-10';
  ws.getRow(9).getCell(4).value = '102'; ws.getRow(9).getCell(5).value = 'Bob';   ws.getRow(9).getCell(6).value = '2026-05-10';
  // BOAT 2 separator at row 13, data at 14-15
  ws.getRow(13).getCell(5).value = 'BAAZ';
  ws.getRow(13).getCell(6).value = 'BOAT 2';
  ws.getRow(14).getCell(4).value = '201'; ws.getRow(14).getCell(5).value = 'Cara'; ws.getRow(14).getCell(6).value = '2026-05-10';
  ws.getRow(15).getCell(4).value = '202'; ws.getRow(15).getCell(5).value = 'Dan';  ws.getRow(15).getCell(6).value = '2026-05-10';
  ws.getRow(18).getCell(6).value = 'TOTAL';

  const layout = {
    headerRow: 6, firstBooking: 7, lastBooking: 17, totalRow: 18, totalMarkerCol: 6,
    layout: { room: 4, guest: 5, departure: 6, adults: 7, child: 8 },
  };
  const result = ctx.inferBoatsFromSheet(ws, layout);
  assert.strictEqual(result.boats.length, 2, `expected 2 boats, got ${result.boats.length}`);
  assert.strictEqual(result.boats[0].name, 'SHADHA');
  assert.strictEqual(result.boats[1].name, 'BAAZ');
  assert.ok(result.boats[0].id && result.boats[0].id.startsWith('b-'), 'boat ids should be prefixed');
  // Each booking row should map to a boat
  const assigned = Object.values(result.assignments);
  assert.strictEqual(assigned.length, 4, `expected 4 assignments, got ${assigned.length}`);
  assert.ok(assigned.filter(id => id === result.boats[0].id).length === 2, 'Alice + Bob → BOAT 1');
  assert.ok(assigned.filter(id => id === result.boats[1].id).length === 2, 'Cara + Dan → BOAT 2');
  console.log('  · separator inference OK');
}

async function testEnsureCapacity(ctx) {
  // 10-row data range (rows 7..16), TOTAL at row 17 with SUM(C7:C16) on col 5 (adults).
  // Style the last data row with a bold font + thick bottom border + yellow fill +
  // number format; assert each inserted row receives that style.
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('test');
  ws.getRow(6).getCell(2).value = 'Room';
  ws.getRow(6).getCell(3).value = 'Guest';
  ws.getRow(6).getCell(4).value = 'Departure';
  ws.getRow(6).getCell(5).value = 'Adults';
  // Style row 16 (last data row) — these styles must clone onto inserted rows.
  const styleSrc = ws.getRow(16);
  for (let c = 2; c <= 5; c++) {
    const cell = styleSrc.getCell(c);
    cell.font = { bold: true };
    cell.border = { bottom: { style: 'thick', color: { argb: 'FF000000' } } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    cell.numFmt = '#,##0.00';
  }
  styleSrc.height = 22;
  // TOTAL row
  ws.getRow(17).getCell(4).value = 'TOTAL';
  ws.getRow(17).getCell(5).value = { formula: 'SUM(E7:E16)', result: 0 };
  // Add a non-standard formula on column 6 to verify it's NOT touched.
  ws.getRow(17).getCell(6).value = { formula: 'SUM(F7:F16)/2', result: 0 };

  const layout = {
    headerRow: 6, firstBooking: 7, lastBooking: 16, totalRow: 17, totalMarkerCol: 4,
    layout: { room: 2, guest: 3, departure: 4, adults: 5 },
  };
  // Need 14 data rows → 4 inserts.
  const warnings = [];
  const newTotal = ctx.ensureCapacity(ws, 14, layout, warnings);
  assert.strictEqual(newTotal, 21, `TOTAL should move to row 21, got ${newTotal}`);
  assert.strictEqual(layout.totalRow, 21, 'layout.totalRow should be mutated');
  // Sum on column 5 (E) should now span E7:E20
  const sumCell = ws.getRow(21).getCell(5).value;
  assert.ok(sumCell && sumCell.formula === 'SUM(E7:E20)', `expected SUM(E7:E20), got ${JSON.stringify(sumCell)}`);
  // Non-standard formula on column 6 should be untouched
  const nonStdCell = ws.getRow(21).getCell(6).value;
  assert.ok(nonStdCell && nonStdCell.formula === 'SUM(F7:F16)/2', `non-standard formula was rewritten: ${JSON.stringify(nonStdCell)}`);
  // Warnings recorded
  assert.ok(warnings.some(w => w.col === 6 && w.why === 'non-standard'),
    'expected a non-standard warning for column 6');
  // Each inserted row carries the style template
  for (let r = 17; r <= 20; r++) {
    const cell = ws.getRow(r).getCell(5);
    assert.ok(cell.font && cell.font.bold === true, `row ${r} col 5 lost bold`);
    assert.ok(cell.fill && cell.fill.fgColor && cell.fill.fgColor.argb === 'FFFFFF00', `row ${r} col 5 lost fill`);
    assert.strictEqual(cell.numFmt, '#,##0.00', `row ${r} col 5 lost numFmt`);
    assert.ok(cell.border && cell.border.bottom && cell.border.bottom.style === 'thick', `row ${r} col 5 lost border`);
  }
  // No-op when requiredRows <= currentRows
  const layout2 = { ...layout, totalRow: 21, lastBooking: 20, layout: layout.layout };
  const same = ctx.ensureCapacity(ws, 5, layout2, []);
  assert.strictEqual(same, 21, 'ensureCapacity should be a no-op when capacity is sufficient');
  console.log('  · ensureCapacity + style cloning + SUM rewrite OK');
}

async function testValidateAssignments(ctx) {
  const ok = ctx.validateAssignments({
    boats: [{ id: 'b-1', name: 'X' }, { id: 'b-2', name: 'Y' }],
    assignments: { '101|Alice|2026-05-10': 'b-1', '102|Bob|2026-05-10': 'b-2' },
  });
  assert.strictEqual(ok.length, 0, `valid input should produce no errors, got ${JSON.stringify(ok)}`);

  const bad = ctx.validateAssignments({
    boats: [{ id: 'b-1', name: 'X' }],
    assignments: { '101|Alice|2026-05-10': 'b-99' },
  });
  assert.strictEqual(bad.length, 1, 'unknown boatId should produce 1 error');
  assert.ok(bad[0].why.includes('unknown boatId'), `error message mismatch: ${bad[0].why}`);
  console.log('  · validateAssignments OK');
}

async function testRebuildBoatGroupings(ctx) {
  // Build a multi-boat sheet, place bookings in scrambled order, run
  // rebuildBoatGroupings, and assert the on-sheet layout is grouped +
  // deterministic across two runs.
  function buildSheet() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('01');
    ws.getRow(6).getCell(4).value = 'Room';
    ws.getRow(6).getCell(5).value = 'Guest';
    ws.getRow(6).getCell(6).value = 'Departure';
    ws.getRow(6).getCell(7).value = 'Adults';
    // Pre-seed 4 bookings on data rows 8..11 in scrambled order.
    ws.getRow(8).getCell(4).value = '101'; ws.getRow(8).getCell(5).value = 'Alice'; ws.getRow(8).getCell(6).value = '2026-05-10'; ws.getRow(8).getCell(7).value = 2;
    ws.getRow(9).getCell(4).value = '102'; ws.getRow(9).getCell(5).value = 'Bob';   ws.getRow(9).getCell(6).value = '2026-05-10'; ws.getRow(9).getCell(7).value = 1;
    ws.getRow(10).getCell(4).value = '201'; ws.getRow(10).getCell(5).value = 'Cara'; ws.getRow(10).getCell(6).value = '2026-05-10'; ws.getRow(10).getCell(7).value = 3;
    ws.getRow(11).getCell(4).value = '202'; ws.getRow(11).getCell(5).value = 'Dan';  ws.getRow(11).getCell(6).value = '2026-05-10'; ws.getRow(11).getCell(7).value = 1;
    ws.getRow(15).getCell(6).value = 'TOTAL';
    return { wb, ws };
  }

  const { wb, ws } = buildSheet();
  const layout = {
    headerRow: 6, firstBooking: 7, lastBooking: 14, totalRow: 15, totalMarkerCol: 6,
    layout: { room: 4, guest: 5, departure: 6, adults: 7, child: 8 },
  };
  const sheetMeta = new Map();
  sheetMeta.set('01', {
    boats: [
      { id: 'b-1', name: 'SHADHA', capacity: 10 },
      { id: 'b-2', name: 'BAAZ',   capacity: 10 },
    ],
    assignments: {
      '101|Alice|2026-05-10': 'b-1',
      '202|Dan|2026-05-10':   'b-1',
      '102|Bob|2026-05-10':   'b-2',
      '201|Cara|2026-05-10':  'b-2',
    },
  });
  const fileEntry = {
    name: '01 - Sunset Dolphin Cruise May.xlsx',
    excursionName: 'Sunset Dolphin Cruise',
    sheets: [{ sheetName: '01', ws, layout }],
    sheetMeta,
    __multiBoat: true,
    dirty: false,
  };

  await ctx.rebuildBoatGroupings(fileEntry);

  // Row 7: separator for SHADHA (BOAT 1)
  assert.strictEqual(ctx.cellStr(ws.getRow(7).getCell(5)), 'SHADHA');
  assert.strictEqual(ctx.cellStr(ws.getRow(7).getCell(6)), 'BOAT 1');
  // Rows 8-9: Alice + Dan in some stable order
  const r8name = ctx.cellStr(ws.getRow(8).getCell(5));
  const r9name = ctx.cellStr(ws.getRow(9).getCell(5));
  assert.ok([r8name, r9name].sort().join(',') === 'Alice,Dan', `BOAT 1 group should be Alice+Dan, got ${r8name}+${r9name}`);
  // Row 10: separator for BAAZ
  assert.strictEqual(ctx.cellStr(ws.getRow(10).getCell(5)), 'BAAZ');
  assert.strictEqual(ctx.cellStr(ws.getRow(10).getCell(6)), 'BOAT 2');
  // Rows 11-12: Bob + Cara
  const r11name = ctx.cellStr(ws.getRow(11).getCell(5));
  const r12name = ctx.cellStr(ws.getRow(12).getCell(5));
  assert.ok([r11name, r12name].sort().join(',') === 'Bob,Cara', `BOAT 2 group should be Bob+Cara, got ${r11name}+${r12name}`);

  // Deterministic: second run produces same layout.
  const { wb: wb2, ws: ws2 } = buildSheet();
  const fileEntry2 = { ...fileEntry, sheets: [{ sheetName: '01', ws: ws2, layout: { ...layout, layout: layout.layout } }] };
  await ctx.rebuildBoatGroupings(fileEntry2);
  for (let r = 7; r <= 12; r++) {
    for (let c = 4; c <= 7; c++) {
      const v1 = String(ctx.cellRaw(ws.getRow(r).getCell(c)) ?? '');
      const v2 = String(ctx.cellRaw(ws2.getRow(r).getCell(c)) ?? '');
      assert.strictEqual(v1, v2, `row ${r} col ${c}: run-to-run mismatch (${v1} vs ${v2})`);
    }
  }
  console.log('  · rebuildBoatGroupings emits ordered separators + deterministic across runs');
}

async function testBulkAssign(ctx) {
  // 3 boats, 6 bookings — all start unassigned. Bulk-assign 3 onto BOAT 1,
  // then the other 3 onto BOAT 2, then unassign one — assert the
  // resulting assignment map and the op-log shape.
  const sheetMeta = new Map();
  sheetMeta.set('01', {
    boats: [
      { id: 'b-1', name: 'SHADHA', capacity: 10 },
      { id: 'b-2', name: 'BAAZ',   capacity: 10 },
      { id: 'b-3', name: 'FALHU',  capacity: 10 },
    ],
    assignments: {},
  });
  const file = {
    name: '01 - Sunset Dolphin Cruise May.xlsx',
    sheets: [{ sheetName: '01' }],
    sheetMeta,
    __multiBoat: true,
    dirty: false,
    pendingOps: [],
  };
  const keysA = ['101|Alice|2026-05-10','102|Bob|2026-05-10','103|Cara|2026-05-10'];
  const keysB = ['201|Dan|2026-05-10','202|Eve|2026-05-10','203|Fay|2026-05-10'];

  const changed1 = ctx.bulkAssignToBoat(file, '01', keysA, 'b-1');
  assert.strictEqual(changed1, 3, `first bulk should change 3, got ${changed1}`);
  const changed2 = ctx.bulkAssignToBoat(file, '01', keysB, 'b-2');
  assert.strictEqual(changed2, 3, `second bulk should change 3, got ${changed2}`);

  let cur = sheetMeta.get('01');
  for (const k of keysA) assert.strictEqual(cur.assignments[k], 'b-1', `${k} should be on b-1`);
  for (const k of keysB) assert.strictEqual(cur.assignments[k], 'b-2', `${k} should be on b-2`);

  // Exactly two setSheetMeta ops should have been recorded.
  assert.strictEqual(file.pendingOps.length, 2,
    `expected 2 setSheetMeta ops, got ${file.pendingOps.length}`);
  for (const op of file.pendingOps) {
    assert.strictEqual(op.kind, 'setSheetMeta');
    assert.ok(op.fields && op.fields.assignments && op.fields.boats,
      'each op should carry both boats and assignments');
  }

  // Unassign two of them in one call (boatId=null).
  const changed3 = ctx.bulkAssignToBoat(file, '01', [keysA[0], keysA[1]], null);
  assert.strictEqual(changed3, 2, `unassign of 2 keys should change 2, got ${changed3}`);
  cur = sheetMeta.get('01');
  assert.ok(!(keysA[0] in cur.assignments), `${keysA[0]} should be unassigned`);
  assert.ok(!(keysA[1] in cur.assignments), `${keysA[1]} should be unassigned`);
  assert.strictEqual(cur.assignments[keysA[2]], 'b-1', `${keysA[2]} should still be on b-1`);

  // Re-applying the same boatId is a no-op (returns 0, no new op recorded).
  const before = file.pendingOps.length;
  const changed4 = ctx.bulkAssignToBoat(file, '01', [keysB[0]], 'b-2');
  assert.strictEqual(changed4, 0, 'reapplying same boat should be a no-op');
  assert.strictEqual(file.pendingOps.length, before, 'no-op must not push a new op');

  // Unknown boatId throws — guards against dangling assignments slipping through.
  assert.throws(() => ctx.bulkAssignToBoat(file, '01', [keysB[0]], 'b-nope'),
    /unknown boatId/, 'unknown boatId should throw');

  console.log('  · bulkAssignToBoat batches, no-ops cleanly, rejects unknown boatId');
}

async function runMultiBoatTests() {
  console.log('\n=== Multi-boat helpers ===');
  console.log('A. Extracting boat helpers from Booking engine v5.html');
  const ctx = loadBoatHelpers();
  // Re-expose cellStr/cellRaw out of the sandbox for assertions in our tests.
  ctx.cellStr = ctx.cellStr || vm.runInContext('cellStr', ctx);
  ctx.cellRaw = ctx.cellRaw || vm.runInContext('cellRaw', ctx);
  console.log('B. inferBoatsFromSheet on a synthesised Sunset-Dolphin sheet');
  await testBoatInference(ctx);
  console.log('C. ensureCapacity preserves style + rewrites simple SUM, leaves non-standard alone');
  await testEnsureCapacity(ctx);
  console.log('D. validateAssignments flags unknown boatIds');
  await testValidateAssignments(ctx);
  console.log('E. rebuildBoatGroupings groups rows + deterministic');
  await testRebuildBoatGroupings(ctx);
  console.log('F. bulkAssignToBoat batches multi-row moves in one op');
  await testBulkAssign(ctx);
  console.log('  All multi-boat assertions passed.');
}

main()
  .then(runMultiBoatTests)
  .catch(e => {
    console.error('\nTEST FAILED:', e && e.stack ? e.stack : e);
    process.exit(1);
  });
