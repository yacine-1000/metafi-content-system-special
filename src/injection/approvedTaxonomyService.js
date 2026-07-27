'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

class ApprovedTaxonomyError extends Error {}
const taxonomyCache = new Map();

function invalidateApprovedTaxonomyCache(workbookPath = null) {
  if (workbookPath) taxonomyCache.delete(path.resolve(workbookPath));
  else taxonomyCache.clear();
}

function cellText(cell) {
  const value = cell.value;
  if (value == null || value === '') return null;
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (value && Array.isArray(value.richText)) return value.richText.map((part) => part.text).join('');
  throw new ApprovedTaxonomyError(`Unsupported taxonomy cell value at ${cell.address}`);
}

function createApprovedTaxonomyService(options = {}) {
  const workbookPath = path.resolve(options.workbookPath || path.join(
    __dirname, '..', '..', 'content', 'script-library', 'source', 'nail_spa_script_library_final.xlsx',
  ));

  async function getTaxonomy() {
    const stat = await fs.promises.stat(workbookPath);
    const version = `${stat.mtimeMs}:${stat.size}`;
    const cached = taxonomyCache.get(workbookPath);
    if (cached && cached.version === version) return cached.value;
    const workbook = new ExcelJS.Workbook();
    try {
      await workbook.xlsx.readFile(workbookPath);
    } catch (error) {
      throw new ApprovedTaxonomyError(`Unable to read approved taxonomy workbook: ${error.message}`);
    }
    const sheet = workbook.getWorksheet('Script Library');
    if (!sheet) throw new ApprovedTaxonomyError('Workbook is missing the Script Library sheet');
    const expectedHeaders = ['Source Set ID', 'Script ID', 'Pillar', 'Pillar Name', 'Version', 'Status'];
    expectedHeaders.forEach((expected, index) => {
      const actual = cellText(sheet.getRow(1).getCell(index + 1));
      if (actual !== expected) {
        throw new ApprovedTaxonomyError(`Script Library column ${index + 1} must be "${expected}"`);
      }
    });

    const pillars = new Set();
    const hookTypes = new Set();
    const formats = new Set(['listicle']);
    const subtopics = [];
    const seenSubtopics = new Set();
    for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      if (!row.hasValues) continue;
      const pillar = cellText(row.getCell(3));
      const pillarName = cellText(row.getCell(4));
      const versionName = cellText(row.getCell(5));
      if (pillar) pillars.add(pillar);
      if (versionName) hookTypes.add(versionName);
      if (pillar && pillarName && !seenSubtopics.has(`${pillar}\0${pillarName}`)) {
        subtopics.push({ subtopic: pillarName, pillar });
        seenSubtopics.add(`${pillar}\0${pillarName}`);
      }
    }
    if (!pillars.size || !hookTypes.size || !subtopics.length) {
      throw new ApprovedTaxonomyError('Approved taxonomy is incomplete');
    }
    const value = {
      pillars: [...pillars],
      subtopics,
      hook_types: [...hookTypes],
      formats: [...formats],
    };
    taxonomyCache.set(workbookPath, { version, value });
    return value;
  }

  return { getTaxonomy };
}

module.exports = { ApprovedTaxonomyError, createApprovedTaxonomyService, invalidateApprovedTaxonomyCache };
