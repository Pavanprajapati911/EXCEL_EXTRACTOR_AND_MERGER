import React from 'react';
import { 
  ChevronUp, 
  ChevronDown, 
  Sparkles, 
  TableProperties, 
  AlertTriangle,
  Heading,
  CheckCircle2,
  Bookmark
} from 'lucide-react';
import { autoDetectHeaderRow, getColLetter } from '../utils/excelProcessor';

export default function HeaderRowPicker({ file, config, onConfigChange, onNotification }) {
  const sheetData = file.sheets[file.activeSheet];
  if (!sheetData) return null;

  const { rawGrid, totalSpreadsheetRows } = sheetData;
  const headerRow = config.headerRow || 1;

  // Header quick modification callbacks
  const handleSetHeaderRow = (newRow) => {
    const parsedRow = parseInt(newRow, 10);
    if (isNaN(parsedRow) || parsedRow < 1) return;
    
    // Automatically shift the extraction startRow to be headerRow + 1
    // to prevent headers from being included in the merged data
    const nextStartRow = Math.min(totalSpreadsheetRows, parsedRow + 1);
    
    onConfigChange(file.id, {
      ...config,
      headerRow: parsedRow,
      startRow: nextStartRow
    });
  };

  const handleIncrement = () => {
    if (headerRow < totalSpreadsheetRows) {
      handleSetHeaderRow(headerRow + 1);
    }
  };

  const handleDecrement = () => {
    if (headerRow > 1) {
      handleSetHeaderRow(headerRow - 1);
    }
  };

  const handleAutoDetect = () => {
    if (!rawGrid || rawGrid.length === 0) return;
    
    const suggestedZeroIdx = autoDetectHeaderRow(rawGrid);
    const suggestedOneIdx = suggestedZeroIdx + 1;
    
    handleSetHeaderRow(suggestedOneIdx);
    onNotification('success', `Intelligent detection suggests Row ${suggestedOneIdx} as headers for "${file.name}".`);
  };

  // Generate slice of rows surrounding the header row for visual grid verification
  const previewSlice = React.useMemo(() => {
    const slice = [];
    const contextRange = 2; // Rows before & after
    
    const startRowIdx = Math.max(0, headerRow - 1 - contextRange);
    const endRowIdx = Math.min(totalSpreadsheetRows - 1, headerRow - 1 + contextRange);

    for (let i = startRowIdx; i <= endRowIdx; i++) {
      const rowValues = rawGrid[i] || [];
      // Pick first 5 columns to keep preview compact
      const displayCells = rowValues.slice(0, 5).map((val, idx) => {
        return val !== undefined && val !== null && String(val).trim() !== "" 
          ? String(val).trim() 
          : "";
      });
      
      let rowType = "data";
      if (i < headerRow - 1) {
        rowType = "metadata";
      } else if (i === headerRow - 1) {
        rowType = "header";
      }

      slice.push({
        excelRowNumber: i + 1,
        cells: displayCells,
        totalCols: rowValues.length,
        type: rowType
      });
    }

    return slice;
  }, [rawGrid, headerRow, totalSpreadsheetRows]);

  // Check if header row is empty
  const isHeaderRowEmpty = React.useMemo(() => {
    const rawRow = rawGrid[headerRow - 1] || [];
    return rawRow.every(val => val === undefined || val === null || String(val).trim() === "");
  }, [rawGrid, headerRow]);

  return (
    <div className="header-picker-panel">
      <div className="picker-toolbar">
        <div className="picker-title-wrap">
          <Heading size={14} className="picker-panel-icon" />
          <span>Column Header Position</span>
        </div>
        
        {/* Row Selector Input and controls */}
        <div className="row-adjuster-controls">
          <button 
            type="button"
            className="btn-picker-adjust"
            onClick={handleDecrement}
            disabled={headerRow <= 1}
            title="Move to previous Excel row"
          >
            <ChevronDown size={14} />
          </button>
          
          <div className="picker-input-wrap">
            <span className="input-prefix">Row</span>
            <input 
              type="number"
              className="picker-num-input"
              value={headerRow}
              min="1"
              max={totalSpreadsheetRows}
              onChange={(e) => handleSetHeaderRow(e.target.value)}
              title="Set row number containing columns names"
            />
          </div>

          <button 
            type="button"
            className="btn-picker-adjust"
            onClick={handleIncrement}
            disabled={headerRow >= totalSpreadsheetRows}
            title="Move to next Excel row"
          >
            <ChevronUp size={14} />
          </button>

          <button 
            type="button"
            className="btn-picker-auto"
            onClick={handleAutoDetect}
            title="Intelligently detect header row by dense text contents"
          >
            <Sparkles size={12} className="btn-icon-inline" />
            Auto-Detect
          </button>
        </div>
      </div>

      {/* Dynamic Warnings */}
      {isHeaderRowEmpty && (
        <div className="picker-warning-banner">
          <AlertTriangle size={12} className="inline-icon" />
          <span><strong>Empty Row Warning:</strong> Selected row {headerRow} contains no text cells! Select a row containing column headers.</span>
        </div>
      )}

      {/* Visual Verification Grid Slice */}
      <div className="picker-grid-verification">
        <div className="grid-verification-title">
          <TableProperties size={12} className="inline-icon text-accent" />
          <span>Context Rows Preview (Around Row {headerRow})</span>
        </div>

        <div className="mini-table-container">
          <table className="mini-preview-table">
            <thead>
              <tr>
                <th className="mini-row-num text-center">Row</th>
                <th className="mini-col-status">Classification</th>
                <th colSpan="5" className="mini-col-data">Cells Content (First 5 Cols)</th>
              </tr>
            </thead>
            <tbody>
              {previewSlice.map(row => {
                const isActive = row.type === 'header';
                return (
                  <tr 
                    key={row.excelRowNumber} 
                    className={`mini-table-row ${isActive ? 'row-header-active' : ''} ${row.type === 'metadata' ? 'row-metadata-muted' : ''}`}
                  >
                    <td className="mini-row-num text-center font-mono">
                      {row.excelRowNumber}
                    </td>
                    <td className="mini-col-status">
                      {row.type === 'header' ? (
                        <span className="badge badge-header">
                          <CheckCircle2 size={10} className="inline-icon" />
                          Column Headers
                        </span>
                      ) : row.type === 'metadata' ? (
                        <span className="badge badge-meta">Metadata / Title</span>
                      ) : (
                        <span className="badge badge-data">Data Row</span>
                      )}
                    </td>
                    {row.cells.map((cellText, cellIdx) => (
                      <td 
                        key={cellIdx} 
                        className={`mini-cell-text ${isActive ? 'header-highlight-cell' : ''} ${cellText === "" ? 'cell-empty' : ''}`}
                        title={cellText || `Column ${getColLetter(cellIdx)}`}
                      >
                        {cellText || "-"}
                      </td>
                    ))}
                    {/* Pad cell columns if less than 5 present */}
                    {row.cells.length < 5 && Array.from({ length: 5 - row.cells.length }).map((_, padIdx) => (
                      <td key={`pad-${padIdx}`} className="mini-cell-text cell-empty">-</td>
                    ))}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
