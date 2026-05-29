import React, { useRef, useState } from 'react';
import { Upload, FileSpreadsheet, Trash2, Layers, MapPin } from 'lucide-react';
import { formatFileSize, parseExcelFile } from '../utils/excelProcessor';
import HeaderRowPicker from './HeaderRowPicker';

export default function FilePicker({ 
  files, 
  rowConfigs, 
  sheetLocations = {},
  uniqueLocations = [],
  onLocationChange,
  onConfigChange, 
  onFilesAdded, 
  onRemoveFile, 
  onSheetChanged, 
  onError,
  onNotification 
}) {
  const [isDragOver, setIsDragOver] = useState(false);
  const [isParsing, setIsParsing] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = () => {
    setIsDragOver(false);
  };

  const processFileList = async (selectedFiles) => {
    setIsParsing(true);
    const parsedFiles = [];
    
    for (let i = 0; i < selectedFiles.length; i++) {
      const file = selectedFiles[i];
      const extension = file.name.split('.').pop().toLowerCase();
      
      if (extension !== 'xlsx' && extension !== 'xls') {
        onError(`"${file.name}" is not a valid Excel file. Only .xlsx and .xls formats are supported.`);
        continue;
      }

      try {
        const parsed = await parseExcelFile(file);
        parsedFiles.push(parsed);
      } catch (err) {
        onError(err.message);
      }
    }

    if (parsedFiles.length > 0) {
      onFilesAdded(parsedFiles);
    }
    setIsParsing(false);
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    setIsDragOver(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      await processFileList(e.dataTransfer.files);
    }
  };

  const handleFileChange = async (e) => {
    if (e.target.files && e.target.files.length > 0) {
      await processFileList(e.target.files);
      e.target.value = ''; // Reset
    }
  };

  const triggerFileSelect = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className="picker-section card">
      <h3 className="section-title">
        <Upload size={18} className="title-icon" />
        Upload Excel Workbooks
      </h3>
      
      {/* Drag & Drop Zone */}
      <div 
        className={`drag-zone ${isDragOver ? 'drag-active' : ''} ${isParsing ? 'drag-disabled' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={isParsing ? null : triggerFileSelect}
        role="button"
        tabIndex={0}
        aria-label="Upload files by dragging here or clicking"
      >
        <input 
          type="file" 
          ref={fileInputRef}
          onChange={handleFileChange}
          multiple 
          accept=".xlsx, .xls"
          className="hidden-file-input"
          style={{ display: 'none' }}
        />
        
        <div className="drag-content">
          <Upload size={42} className="upload-cloud-icon" />
          {isParsing ? (
            <p className="drag-text font-medium">Parsing workbooks, please wait...</p>
          ) : (
            <>
              <p className="drag-text font-semibold">
                Drag & drop your Excel files here, or <span className="highlight-text">browse</span>
              </p>
              <p className="drag-subtext">Supports .xlsx and .xls formats</p>
            </>
          )}
        </div>
      </div>

      {/* File List */}
      {files.length > 0 && (
        <div className="file-list-container">
          <div className="file-list-header">
            <span>Uploaded Files & Header Setup ({files.length})</span>
          </div>
          
          <ul className="file-list" aria-label="Uploaded Excel Files">
            {files.map((file) => {
              const activeSheetData = file.sheets[file.activeSheet];
              const totalRows = activeSheetData ? activeSheetData.totalSpreadsheetRows : 0;
              
              return (
                <li key={file.id} className="file-item-wrapper">
                  {/* File Metadata Header Bar */}
                  <div className="file-item-header">
                    <div className="file-info-col">
                      <FileSpreadsheet className="excel-file-icon" size={22} />
                      <div className="file-meta">
                        <span className="file-name" title={file.name}>{file.name}</span>
                        <span className="file-details">
                          {formatFileSize(file.size)} | {file.sheetNames.length} sheet(s)
                        </span>
                      </div>
                    </div>

                    <div className="file-actions-col">
                      {/* Worksheet Selection Dropdown */}
                      {file.sheetNames.length > 1 ? (
                        <div className="sheet-selector-wrapper">
                          <Layers size={14} className="selector-icon" />
                          <select 
                            value={file.activeSheet}
                            onChange={(e) => onSheetChanged(file.id, e.target.value)}
                            className="sheet-dropdown"
                            title="Select Excel Worksheet"
                          >
                            {file.sheetNames.map((sheetName) => (
                              <option key={sheetName} value={sheetName}>
                                {sheetName} ({file.sheets[sheetName]?.totalSpreadsheetRows} rows)
                              </option>
                            ))}
                          </select>
                        </div>
                      ) : (
                        <span className="single-sheet-badge">
                          Sheet: <strong className="badge-name">{file.activeSheet}</strong> ({totalRows} rows)
                        </span>
                      )}

                      {/* Remove file button */}
                      <button 
                        type="button"
                        onClick={() => onRemoveFile(file.id)}
                        className="btn-icon btn-danger"
                        title="Remove File"
                        aria-label={`Remove ${file.name}`}
                      >
                        <Trash2 size={16} />
                      </button>
                    </div>
                  </div>

                  {/* Elegant Location Assignment Section */}
                  <div className="file-location-section">
                    <div className="location-input-wrapper">
                      <MapPin size={14} className="location-icon" />
                      <span className="location-label">Location:</span>
                      <input 
                        type="text"
                        list={`location-suggestions-${file.id}`}
                        value={sheetLocations[file.id]?.[file.activeSheet] || ""}
                        onChange={(e) => onLocationChange(file.id, file.activeSheet, e.target.value)}
                        placeholder="Type or select location for this sheet..."
                        className="location-textbox"
                        title="Location assigned to this worksheet"
                      />
                      <datalist id={`location-suggestions-${file.id}`}>
                        {uniqueLocations.map(loc => (
                          <option key={loc} value={loc} />
                        ))}
                      </datalist>
                    </div>
                  </div>

                  {/* Configurable Header Row Selector Panel */}
                  {rowConfigs[file.id] && (
                    <HeaderRowPicker 
                      file={file}
                      config={rowConfigs[file.id]}
                      onConfigChange={onConfigChange}
                      onNotification={onNotification}
                    />
                  )}
                </li>
              );
            })}
          </ul>
        </div>
      )}
    </div>
  );
}
