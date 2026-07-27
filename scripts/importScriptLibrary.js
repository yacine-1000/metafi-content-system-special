'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const ROOT = path.resolve(__dirname, '..');
const LIBRARY_DIR = path.join(ROOT, 'content', 'script-library');
const WORKBOOK_PATH = path.join(LIBRARY_DIR, 'source', 'nail_spa_script_library_final.xlsx');
const SOURCE_SETS_DIR = path.join(LIBRARY_DIR, 'source-sets');
const INDEX_PATH = path.join(LIBRARY_DIR, 'index.json');
const SHEET_NAME = 'Script Library';
const SLIDE_TWO_TEXT = 'احفظي المقطع\nبتحتاجينه';
const EXPECTED_HEADERS = [
  'Source Set ID', 'Script ID', 'Pillar', 'Pillar Name', 'Version', 'Status',
  'Slide 1', 'Slide 2', 'Slide 3', 'Slide 4', 'Slide 5', 'Slide 6',
  'Slide 7', 'Slide 8', 'Slide 9', 'Slide 10', 'Slide 11', 'Slide 12',
];
const SOURCE_SET_ID_PATTERN = /^SET-\d{3,}$/;
const PILLAR_PATTERN = /^P[1-4]$/;

class ScriptLibraryImportError extends Error {
  constructor(message, report = null) {
    super(message);
    this.name = 'ScriptLibraryImportError';
    this.report = report;
  }
}

function textValue(cell, location, { required = false } = {}) {
  const value = cell.value;
  if (value == null || value === '') {
    if (required) throw new ScriptLibraryImportError(`${location} is required`);
    return null;
  }
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value && Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('');
  if (value && Object.prototype.hasOwnProperty.call(value, 'formula')) {
    throw new ScriptLibraryImportError(`${location} must contain text, not a formula`);
  }
  throw new ScriptLibraryImportError(`${location} contains an unsupported Excel value`);
}

function writeJsonAtomic(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const temporaryPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(temporaryPath, `${JSON.stringify(value, null, 2)}\n`, 'utf8');
  fs.renameSync(temporaryPath, filePath);
}

function validateHeaders(sheet) {
  if (!sheet) throw new ScriptLibraryImportError(`Workbook is missing the "${SHEET_NAME}" sheet`);
  const actual = EXPECTED_HEADERS.map((_header, index) => textValue(
    sheet.getRow(1).getCell(index + 1),
    `${SHEET_NAME}!${sheet.getRow(1).getCell(index + 1).address}`,
  ));
  EXPECTED_HEADERS.forEach((expected, index) => {
    if (actual[index] !== expected) {
      throw new ScriptLibraryImportError(
        `${SHEET_NAME} column ${index + 1} must be "${expected}"; found "${actual[index] || ''}"`,
      );
    }
  });
}

function parseAndValidate(sheet) {
  const errors = [];
  const warnings = [];
  const scriptIds = new Set();
  const duplicateScriptIds = new Set();
  const sourceSets = new Map();
  const scriptsPerPillar = { P1: 0, P2: 0, P3: 0, P4: 0 };
  let scriptCount = 0;

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const row = sheet.getRow(rowNumber);
    if (!row.hasValues) continue;
    const at = (column) => `${SHEET_NAME}!${row.getCell(column).address}`;
    const readRequired = (column, label) => {
      try {
        return textValue(row.getCell(column), at(column), { required: true });
      } catch (error) {
        errors.push(`${label}: ${error.message}`);
        return null;
      }
    };
    const sourceSetId = readRequired(1, 'missing Source Set ID');
    const scriptId = readRequired(2, 'missing Script ID');
    const pillar = readRequired(3, 'missing Pillar');
    const pillarName = readRequired(4, 'missing Pillar Name');
    const version = readRequired(5, 'missing Version');
    const status = readRequired(6, 'missing Status');

    if (!sourceSetId || !scriptId || !pillar || !pillarName || !version || !status) continue;
    scriptCount += 1;
    if (!SOURCE_SET_ID_PATTERN.test(sourceSetId)) errors.push(`${at(1)} has invalid Source Set ID "${sourceSetId}"`);
    if (!PILLAR_PATTERN.test(pillar)) errors.push(`${at(3)} must be P1, P2, P3, or P4`);
    if (scriptIds.has(scriptId)) duplicateScriptIds.add(scriptId);
    scriptIds.add(scriptId);
    if (PILLAR_PATTERN.test(pillar)) scriptsPerPillar[pillar] += 1;

    const slideValues = [];
    for (let slideNumber = 1; slideNumber <= 12; slideNumber += 1) {
      try {
        slideValues.push(textValue(row.getCell(6 + slideNumber), at(6 + slideNumber)));
      } catch (error) {
        errors.push(error.message);
        slideValues.push(null);
      }
    }
    const populatedNumbers = slideValues
      .map((text, index) => (text == null ? null : index + 1))
      .filter((number) => number != null);
    const lastSlideNumber = populatedNumbers.length ? populatedNumbers[populatedNumbers.length - 1] : 0;

    if (slideValues[0] == null) errors.push(`${at(7)} is missing Slide 1`);
    if (slideValues[1] !== SLIDE_TWO_TEXT) errors.push(`${at(8)} has incorrect Slide 2`);
    if (lastSlideNumber < 3 || lastSlideNumber > 12) {
      errors.push(`${at(7)}:${at(18)} has ${lastSlideNumber} slides; expected 3 through 12`);
    }
    for (let slideNumber = 1; slideNumber <= lastSlideNumber; slideNumber += 1) {
      if (slideValues[slideNumber - 1] == null) {
        errors.push(`${at(6 + slideNumber)} is a blank gap before Slide ${lastSlideNumber}`);
      }
    }
    if (!lastSlideNumber || slideValues[lastSlideNumber - 1] == null) {
      errors.push(`${at(6 + Math.max(lastSlideNumber, 1))} has an empty final CTA`);
    }

    const slides = [];
    for (let slideNumber = 1; slideNumber <= lastSlideNumber; slideNumber += 1) {
      const text = slideValues[slideNumber - 1];
      if (text == null) continue;
      slides.push({
        slide_number: slideNumber,
        slide_label: `Slide ${slideNumber}`,
        is_metafi_slide: slideNumber === lastSlideNumber,
        text,
      });
    }

    const metadata = {
      source_set_id: sourceSetId,
      pillar,
      pillar_name: pillarName,
      subtopic: pillarName,
      topic: pillarName,
    };
    if (!sourceSets.has(sourceSetId)) sourceSets.set(sourceSetId, { ...metadata, scripts: [] });
    const sourceSet = sourceSets.get(sourceSetId);
    for (const field of ['pillar', 'pillar_name']) {
      if (sourceSet[field] !== metadata[field]) {
        errors.push(`${at(1)} conflicts with ${field} already stored for ${sourceSetId}`);
      }
    }
    sourceSet.scripts.push({
      script_id: scriptId,
      script_version: version,
      status,
      hook_type: version,
      format: 'listicle',
      original_slide_count: lastSlideNumber,
      final_slide_count: lastSlideNumber,
      slides,
    });
  }

  for (const scriptId of duplicateScriptIds) errors.push(`duplicate Script ID "${scriptId}"`);
  if (!sourceSets.size) errors.push(`${SHEET_NAME} contains no Source Sets`);
  const report = {
    source_set_count: sourceSets.size,
    script_count: scriptCount,
    scripts_per_pillar: scriptsPerPillar,
    duplicate_script_ids: [...duplicateScriptIds].sort(),
    missing_source_set_ids: errors.filter((error) => error.startsWith('missing Source Set ID')).length,
    missing_slide_1: errors.filter((error) => error.endsWith('is missing Slide 1')).length,
    incorrect_slide_2: errors.filter((error) => error.endsWith('has incorrect Slide 2')).length,
    invalid_slide_counts: errors.filter((error) => error.includes('slides; expected 3 through 12')).length,
    blank_slide_gaps: errors.filter((error) => error.includes('is a blank gap before Slide')).length,
    empty_final_cta: errors.filter((error) => error.endsWith('has an empty final CTA')).length,
    warnings,
    errors,
  };
  return { sourceSets, report };
}

async function importScriptLibrary() {
  if (!fs.existsSync(WORKBOOK_PATH)) {
    throw new ScriptLibraryImportError(`Workbook is missing: ${path.relative(ROOT, WORKBOOK_PATH)}`);
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(WORKBOOK_PATH);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  validateHeaders(sheet);
  const { sourceSets, report } = parseAndValidate(sheet);
  if (report.errors.length) {
    throw new ScriptLibraryImportError(`Validation failed with ${report.errors.length} error(s)`, report);
  }

  const ordered = Array.from(sourceSets.values()).sort((left, right) => (
    left.source_set_id.localeCompare(right.source_set_id, undefined, { numeric: true })
  ));
  fs.mkdirSync(SOURCE_SETS_DIR, { recursive: true });
  for (const sourceSet of ordered) {
    writeJsonAtomic(path.join(SOURCE_SETS_DIR, `${sourceSet.source_set_id}.json`), sourceSet);
  }
  const expectedFiles = new Set(ordered.map((sourceSet) => `${sourceSet.source_set_id}.json`));
  for (const filename of fs.readdirSync(SOURCE_SETS_DIR)) {
    if (path.extname(filename).toLowerCase() === '.json' && !expectedFiles.has(filename)) {
      fs.unlinkSync(path.join(SOURCE_SETS_DIR, filename));
    }
  }
  const index = {
    source_sets: ordered.map((sourceSet) => ({
      source_set_id: sourceSet.source_set_id,
      file: `source-sets/${sourceSet.source_set_id}.json`,
      pillar: sourceSet.pillar,
      pillar_name: sourceSet.pillar_name,
      subtopic: sourceSet.subtopic,
      topic: sourceSet.topic,
      hook_types: Array.from(new Set(sourceSet.scripts.map((script) => script.hook_type))),
      script_count: sourceSet.scripts.length,
    })),
  };
  writeJsonAtomic(INDEX_PATH, index);
  return report;
}

if (require.main === module) {
  importScriptLibrary()
    .then((report) => console.log(JSON.stringify(report, null, 2)))
    .catch((error) => {
      console.error(`Script Library import failed: ${error.message}`);
      if (error.report) console.error(JSON.stringify(error.report, null, 2));
      process.exitCode = 1;
    });
}

module.exports = { ScriptLibraryImportError, importScriptLibrary };
