import * as XLSX from 'xlsx';

/**
 * Utility to convert column index to Excel column letter (0 -> A, 27 -> AB)
 */
export function getColLetter(colIdx) {
  let temp, letter = "";
  while (colIdx >= 0) {
    temp = colIdx % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    colIdx = Math.floor(colIdx / 26) - 1;
  }
  return letter;
}

/**
 * Formats file size in bytes into human readable string
 */
export function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Extract columns/headers dynamically from a specific row in the raw 2D grid
 * @param {Array} rawGrid 
 * @param {number} headerRowIdx 0-indexed row index
 * @returns {Array} List of header objects { key, letter, index }
 */
export function extractHeadersFromRow(rawGrid, headerRowIdx) {
  if (!rawGrid || headerRowIdx < 0 || headerRowIdx >= rawGrid.length) {
    return [];
  }
  const rawRow = rawGrid[headerRowIdx] || [];
  const usedNames = {};
  
  return rawRow.map((val, colIdx) => {
    let nameText = "";
    if (val !== undefined && val !== null && String(val).trim() !== "") {
      nameText = String(val).trim();
    } else {
      nameText = `Column ${getColLetter(colIdx)}`;
    }
    
    // Resolve duplicate column names in the same sheet
    if (usedNames[nameText] !== undefined) {
      usedNames[nameText]++;
      nameText = `${nameText} (${usedNames[nameText]})`;
    } else {
      usedNames[nameText] = 1;
    }
    
    return {
      key: nameText,
      letter: getColLetter(colIdx),
      index: colIdx
    };
  });
}

/**
 * Heuristically auto-detect which row contains column names (column headers)
 * Evaluates dense cell population and text-to-number distributions in the first 25 rows
 * @param {Array} rawGrid 
 * @returns {number} The 0-indexed header row suggestion
 */
export function autoDetectHeaderRow(rawGrid) {
  if (!rawGrid || rawGrid.length === 0) return 0;
  
  const maxScanRows = Math.min(rawGrid.length, 25);
  let bestRowIdx = 0;
  let highestScore = -1;
  
  for (let r = 0; r < maxScanRows; r++) {
    const row = rawGrid[r] || [];
    let nonEmptyCount = 0;
    let numericCount = 0;
    
    row.forEach(val => {
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        nonEmptyCount++;
        if (typeof val === 'number') {
          numericCount++;
        }
      }
    });
    
    // If a row is purely numbers, it's likely a data row, not headers
    if (nonEmptyCount > 0 && numericCount === nonEmptyCount) {
      continue;
    }
    
    // Header score: higher count of non-empty string headers is preferred
    const score = nonEmptyCount - (numericCount * 0.5);
    if (score > highestScore) {
      highestScore = score;
      bestRowIdx = r;
    }
  }
  
  return bestRowIdx;
}

/**
 * Parse uploaded Excel file and extract raw 2D grid structures per sheet.
 * @param {File} file 
 * @returns {Promise<Object>} Parsed file configuration object containing raw grids
 */
export function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: false, cellText: false });
        
        const sheetNames = workbook.SheetNames;
        if (!sheetNames || sheetNames.length === 0) {
          throw new Error("No worksheets found in this Excel file.");
        }

        const sheets = {};

        sheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          // Convert sheets directly into raw 2D grids (array of arrays)
          const rawGrid = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
          
          sheets[sheetName] = {
            rawGrid,
            totalSpreadsheetRows: rawGrid.length
          };
        });

        resolve({
          id: Math.random().toString(36).substring(2, 9),
          name: file.name,
          size: file.size,
          sheetNames,
          activeSheet: sheetNames[0], // Default to first sheet
          sheets
        });
      } catch (error) {
        reject(new Error(`Failed to parse "${file.name}": ${error.message}`));
      }
    };

    reader.onerror = () => {
      reject(new Error(`Error reading file "${file.name}"`));
    };

    reader.readAsArrayBuffer(file);
  });
}

/**
 * Normalizes rows and merges datasets dynamically from raw grids based on custom rowConfigs and columnMappings
 * @param {Array} files List of parsed file configurations
 * @param {Array} selectedColumns List of selected column names to merge
 * @param {Object} rowConfigs Map of file ID to configuration { headerRow, startRow, endRow, skipEmptyRows }
 * @param {Object} columnMappings Dictionary of custom merged target columns { [targetName]: { [fileId]: originalName } }
 * @returns {Array} List of combined row objects
 */
export function mergeDatasets(files, selectedColumns, rowConfigs, columnMappings = {}, sheetLocations = {}) {
  const mergedRows = [];

  files.forEach(file => {
    const config = rowConfigs[file.id] || { headerRow: 1, startRow: 2, endRow: 999999, skipEmptyRows: true };
    const sheetData = file.sheets[file.activeSheet];
    if (!sheetData) return;

    const { rawGrid } = sheetData;
    
    // Extract headers dynamically for this specific file sheet based on its headerRow configuration
    const headerRowIdx = Math.max(0, config.headerRow - 1);
    const headers = extractHeadersFromRow(rawGrid, headerRowIdx);

    // Find column indexes in this specific file sheet for selected columns
    const colMap = {};
    selectedColumns.forEach(colName => {
      // Check if this column maps to a unified target column
      const originalColName = columnMappings[colName]?.[file.id] || colName;
      const headerItem = headers.find(h => h.key === originalColName);
      colMap[colName] = headerItem ? headerItem.index : -1;
    });

    // Excel is 1-indexed. Row 1 matches rawGrid[0].
    // Extraction start index is config.startRow - 1.
    const startIdx = Math.max(0, config.startRow - 1);
    const endIdx = config.endRow ? Math.min(rawGrid.length, config.endRow) : rawGrid.length;

    // Ensure we do not pull the header row itself or rows prior to it by mistake
    const trueStartIdx = Math.max(startIdx, headerRowIdx + 1);

    for (let i = trueStartIdx; i < endIdx; i++) {
      const row = rawGrid[i];
      if (!row) continue;

      // Check if row is empty across selected columns
      const isRowEmpty = row.every(val => val === undefined || val === null || String(val).trim() === "");
      if (config.skipEmptyRows && isRowEmpty) {
        continue;
      }

      // Create a normalized row object with unified columns
      const normalizedRow = {
        _sourceFile: file.name,
        _sourceSheet: file.activeSheet,
        _sourceExcelRow: i + 1 // The 1-based original row in the Excel sheet
      };

      selectedColumns.forEach(colName => {
        if (colName === 'Location') {
          const assignedLoc = sheetLocations[file.id]?.[file.activeSheet];
          if (assignedLoc !== undefined && assignedLoc.trim() !== "") {
            normalizedRow[colName] = assignedLoc;
            return;
          }
        }

        const colIdx = colMap[colName];
        if (colIdx !== -1 && row[colIdx] !== undefined) {
          const val = row[colIdx];
          if (val instanceof Date) {
            normalizedRow[colName] = val;
          } else if (typeof val === 'number') {
            normalizedRow[colName] = val;
          } else {
            normalizedRow[colName] = String(val);
          }
        } else {
          normalizedRow[colName] = ""; // Pad missing column values
        }
      });

      mergedRows.push(normalizedRow);
    }
  });

  return mergedRows;
}

/**
 * Generate a new Excel workbook containing merged rows and trigger download.
 * @param {Array} data The merged JSON dataset
 * @param {Array} columns List of columns to export
 * @param {string} fileName Custom output filename
 */
export function exportMergedToExcel(data, columns, fileName) {
  if (!data || data.length === 0) {
    throw new Error("No data to export.");
  }

  // Format data for standard export by removing metadata audit columns
  const exportRows = data.map(row => {
    const cleanRow = {};
    columns.forEach(col => {
      cleanRow[col] = row[col];
    });
    return cleanRow;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportRows, { header: columns });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Merged Data");

  let finalName = fileName.trim() || "merged_dataset";
  if (!finalName.endsWith(".xlsx")) {
    finalName += ".xlsx";
  }

  XLSX.writeFile(workbook, finalName);
}
