import React from 'react';
import { Sliders, HelpCircle, Check, AlertTriangle } from 'lucide-react';

export default function RowRangeSelector({ files, rowConfigs, onConfigChange }) {

  const handleInputChange = (fileId, field, val) => {
    const fileConfig = rowConfigs[fileId] || { startRow: 2, endRow: 999999, skipEmptyRows: true };
    const numVal = parseInt(val, 10);
    
    const newConfig = {
      ...fileConfig,
      [field]: isNaN(numVal) ? val : numVal // Allow raw text input while typing
    };
    onConfigChange(fileId, newConfig);
  };

  const handleCheckboxChange = (fileId, val) => {
    const fileConfig = rowConfigs[fileId] || { startRow: 2, endRow: 999999, skipEmptyRows: true };
    const newConfig = {
      ...fileConfig,
      skipEmptyRows: val
    };
    onConfigChange(fileId, newConfig);
  };

  if (files.length === 0) {
    return null;
  }

  return (
    <div className="range-section card">
      <h3 className="section-title">
        <Sliders size={18} className="title-icon" />
        Configure Row Ranges (per File)
      </h3>
      
      <p className="range-section-helper">
        By default, the first row (Spreadsheet Row 1) is assumed to contain headers. Adjust start and end index parameters below if needed.
      </p>

      <div className="file-range-grid" aria-label="Configure row range settings per file">
        {files.map(file => {
          const config = rowConfigs[file.id] || { startRow: 2, endRow: 999999, skipEmptyRows: true };
          const sheetData = file.sheets[file.activeSheet];
          // Total rows in spreadsheet (including header, raw rows starts at index 1 of Javascript array, raw grid has totalSpreadsheetRows)
          const totalSpreadsheetRows = sheetData ? sheetData.totalSpreadsheetRows : 0;
          const maxAvailableDataRow = totalSpreadsheetRows;

          // Set default endRow if not customized yet
          const currentEndRow = config.endRow === 999999 ? maxAvailableDataRow : config.endRow;

          // Validation
          const startVal = parseInt(config.startRow, 10);
          const endVal = parseInt(currentEndRow, 10);
          const isStartInvalid = isNaN(startVal) || startVal < 1;
          const isEndInvalid = isNaN(endVal) || endVal < 1 || (startVal && endVal < startVal);
          const hasError = isStartInvalid || isEndInvalid;

          return (
            <div key={file.id} className={`file-range-card ${hasError ? 'card-error' : ''}`}>
              <div className="range-card-header">
                <span className="range-card-filename" title={file.name}>{file.name}</span>
                <span className="range-card-sheetname">
                  Worksheet: <strong>{file.activeSheet}</strong>
                </span>
              </div>

              <div className="range-card-body">
                {/* Inputs Row */}
                <div className="range-inputs-wrapper">
                  <div className="range-input-group">
                    <label className="range-input-label">Start Row</label>
                    <input 
                      type="number"
                      min="1"
                      className={`range-num-input ${isStartInvalid ? 'input-error' : ''}`}
                      value={config.startRow}
                      onChange={(e) => handleInputChange(file.id, 'startRow', e.target.value)}
                      title="Start row (1-indexed)"
                    />
                  </div>

                  <div className="range-input-group">
                    <label className="range-input-label">End Row</label>
                    <input 
                      type="number"
                      min="1"
                      className={`range-num-input ${isEndInvalid ? 'input-error' : ''}`}
                      value={currentEndRow}
                      onChange={(e) => handleInputChange(file.id, 'endRow', e.target.value)}
                      placeholder={maxAvailableDataRow}
                      title="End row (inclusive, 1-indexed)"
                    />
                  </div>

                  <div className="range-checkbox-group">
                    <label className="checkbox-toggle-container">
                      <input 
                        type="checkbox"
                        checked={config.skipEmptyRows}
                        onChange={(e) => handleCheckboxChange(file.id, e.target.checked)}
                      />
                      <span className="toggle-label-text">Skip empty rows</span>
                    </label>
                  </div>
                </div>

                {/* Validation Warnings / Status messages */}
                <div className="range-validation-status">
                  {isStartInvalid && (
                    <span className="validation-msg error">
                      <AlertTriangle size={12} className="inline-icon" />
                      Start row must be a positive integer &ge; 1.
                    </span>
                  )}
                  {!isStartInvalid && isEndInvalid && (
                    <span className="validation-msg error">
                      <AlertTriangle size={12} className="inline-icon" />
                      End row must be &ge; Start row.
                    </span>
                  )}
                  {!hasError && (
                    <span className="validation-msg success">
                      <Check size={12} className="inline-icon animate-pulse" />
                      Extracting from spreadsheet row <strong>{startVal}</strong> to <strong>{endVal}</strong> (Max {totalSpreadsheetRows} rows in file)
                    </span>
                  )}
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
