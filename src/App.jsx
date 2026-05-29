import React, { useState, useMemo, useEffect } from 'react';
import { 
  GitMerge, 
  Settings, 
  RefreshCw, 
  Check, 
  AlertCircle, 
  FileDown,
  Sparkles
} from 'lucide-react';
import FilePicker from './components/FilePicker';
import ColumnMapper from './components/ColumnMapper';
import RowRangeSelector from './components/RowRangeSelector';
import PreviewTable from './components/PreviewTable';
import { mergeDatasets, exportMergedToExcel, autoDetectHeaderRow } from './utils/excelProcessor';
import { saveWorkspace, loadWorkspace, clearWorkspace } from './utils/db';

export default function App() {
  const [files, setFiles] = useState([]);
  const [rowConfigs, setRowConfigs] = useState({});
  const [columnMappings, setColumnMappings] = useState({});
  const [sheetLocations, setSheetLocations] = useState({});
  const [outputFileName, setOutputFileName] = useState('merged_excel_output');
  const [notification, setNotification] = useState(null);
  const [isExporting, setIsExporting] = useState(false);
  const [isInitialLoad, setIsInitialLoad] = useState(true);

  // Clear notifications automatically
  useEffect(() => {
    if (notification) {
      const timer = setTimeout(() => setNotification(null), 6000);
      return () => clearTimeout(timer);
    }
  }, [notification]);

  // Toast / Alert Trigger
  const triggerNotification = (type, message) => {
    setNotification({ type, message });
  };

  // 1. Initial workspace loading from IndexedDB on mount
  useEffect(() => {
    async function initWorkspace() {
      try {
        const savedState = await loadWorkspace();
        if (savedState) {
          // Populate states
          setFiles(savedState.files || []);
          setRowConfigs(savedState.rowConfigs || {});
          setColumnMappings(savedState.columnMappings || {});
          setSheetLocations(savedState.sheetLocations || {});
          if (savedState.outputFileName) {
            setOutputFileName(savedState.outputFileName);
          }
          triggerNotification('success', 'Restored previous workspace from browser database.');
        }
      } catch (err) {
        console.error("Workspace restoration error:", err);
      } finally {
        setIsInitialLoad(false);
      }
    }
    initWorkspace();
  }, []);

  // 2. Background Auto-Saving Hook to IndexedDB
  useEffect(() => {
    if (isInitialLoad) return;

    // Trigger asynchronous save in background
    saveWorkspace(files, rowConfigs, columnMappings, outputFileName, sheetLocations);
  }, [files, rowConfigs, columnMappings, outputFileName, sheetLocations, isInitialLoad]);

  // Unique list of all locations entered across sheets
  const uniqueLocations = useMemo(() => {
    const set = new Set();
    Object.values(sheetLocations).forEach(fileLocs => {
      if (fileLocs) {
        Object.values(fileLocs).forEach(loc => {
          if (loc && loc.trim() !== "") {
            set.add(loc.trim());
          }
        });
      }
    });
    return Array.from(set).sort();
  }, [sheetLocations]);

  // 3. Derive output columns list directly as the set of defined target aliases in the mapping studio
  const selectedColumns = useMemo(() => {
    const cols = Object.keys(columnMappings);
    
    // Check if any sheet has a location assigned
    const hasAnyLocation = files.some(file => {
      const locs = sheetLocations[file.id];
      return locs && Object.values(locs).some(val => val && val.trim() !== "");
    });

    if (hasAnyLocation && !cols.includes('Location')) {
      return [...cols, 'Location'];
    }
    return cols;
  }, [columnMappings, files, sheetLocations]);

  // 4. Row range and header validation checks
  const rangeValidation = useMemo(() => {
    let isValid = true;
    let errorMsg = "";

    files.forEach(file => {
      const config = rowConfigs[file.id];
      if (config) {
        const header = parseInt(config.headerRow, 10);
        const start = parseInt(config.startRow, 10);
        const end = parseInt(config.endRow, 10);
        
        if (isNaN(header) || header < 1) {
          isValid = false;
          errorMsg = `Header row must be a valid integer >= 1 for "${file.name}"`;
        } else if (isNaN(start) || start < 1) {
          isValid = false;
          errorMsg = `Start row must be a valid integer >= 1 for "${file.name}"`;
        } else if (start <= header) {
          isValid = false;
          errorMsg = `Start row must be greater than Header row (${header}) for "${file.name}" to exclude headers from data.`;
        } else if (isNaN(end) || end < start) {
          isValid = false;
          errorMsg = `End row must be greater than or equal to Start row for "${file.name}"`;
        }
      }
    });

    return { isValid, errorMsg };
  }, [files, rowConfigs]);

  // 5. Merging Core Dataset with custom Column Mappings dictionary
  const mergedData = useMemo(() => {
    if (files.length === 0 || selectedColumns.length === 0 || !rangeValidation.isValid) {
      return [];
    }
    try {
      return mergeDatasets(files, selectedColumns, rowConfigs, columnMappings, sheetLocations);
    } catch (err) {
      console.error("Merging error:", err);
      return [];
    }
  }, [files, selectedColumns, rowConfigs, rangeValidation, columnMappings, sheetLocations]);

  // Action: Add Uploaded Files
  const handleFilesAdded = (newFiles) => {
    setFiles(prev => {
      const updated = [...prev, ...newFiles];
      
      const newConfigs = { ...rowConfigs };
      newFiles.forEach(file => {
        const activeSheetData = file.sheets[file.activeSheet];
        
        // Auto-detect suggestions on upload
        const suggestedZeroIdx = activeSheetData ? autoDetectHeaderRow(activeSheetData.rawGrid) : 0;
        const suggestedHeaderRow = suggestedZeroIdx + 1;
        const maxRows = activeSheetData ? activeSheetData.totalSpreadsheetRows : 0;
        
        newConfigs[file.id] = {
          headerRow: suggestedHeaderRow,
          startRow: Math.min(maxRows, suggestedHeaderRow + 1),
          endRow: maxRows,
          skipEmptyRows: true
        };
      });
      setRowConfigs(newConfigs);

      return updated;
    });

    triggerNotification('success', `Loaded ${newFiles.length} new workbook(s). Heuristics automatically suggested custom column rows.`);
  };

  // Action: Remove specific file
  const handleRemoveFile = (fileId) => {
    const fileToRemove = files.find(f => f.id === fileId);
    
    // Clean up mapping records corresponding to the deleted file
    setColumnMappings(prev => {
      const updated = {};
      Object.keys(prev).forEach(target => {
        const fileMap = { ...prev[target] };
        delete fileMap[fileId];
        // Retain alias rules if at least one file mapping remains
        if (Object.keys(fileMap).length > 0) {
          updated[target] = fileMap;
        }
      });
      return updated;
    });

    setFiles(prev => {
      const updated = prev.filter(f => f.id !== fileId);
      
      const newConfigs = { ...rowConfigs };
      delete newConfigs[fileId];
      setRowConfigs(newConfigs);

      return updated;
    });

    // Clean up locations mapping
    setSheetLocations(prev => {
      const updated = { ...prev };
      delete updated[fileId];
      return updated;
    });

    if (fileToRemove) {
      triggerNotification('info', `Removed "${fileToRemove.name}"`);
    }
  };

  // Action: Sheet changed per file
  const handleSheetChanged = (fileId, newSheet) => {
    setFiles(prev => {
      const updated = prev.map(f => {
        if (f.id === fileId) {
          return { ...f, activeSheet: newSheet };
        }
        return f;
      });

      const file = updated.find(f => f.id === fileId);
      if (file) {
        const activeSheetData = file.sheets[newSheet];
        
        const suggestedZeroIdx = activeSheetData ? autoDetectHeaderRow(activeSheetData.rawGrid) : 0;
        const suggestedHeaderRow = suggestedZeroIdx + 1;
        const maxRows = activeSheetData ? activeSheetData.totalSpreadsheetRows : 0;
        
        setRowConfigs(prevConfigs => ({
          ...prevConfigs,
          [fileId]: {
            ...prevConfigs[fileId],
            headerRow: suggestedHeaderRow,
            startRow: Math.min(maxRows, suggestedHeaderRow + 1),
            endRow: maxRows
          }
        }));
      }

      return updated;
    });

    // Reset mapping records for this specific file on sheet switch
    setColumnMappings(prev => {
      const updated = {};
      Object.keys(prev).forEach(target => {
        const fileMap = { ...prev[target] };
        delete fileMap[fileId];
        if (Object.keys(fileMap).length > 0) {
          updated[target] = fileMap;
        }
      });
      return updated;
    });

    triggerNotification('info', 'Active sheet switched. Resetting aliases mapping for this file.');
  };

  // Column Mapping Workspace callbacks
  const handleAddMapping = (targetName, selections) => {
    setColumnMappings(prev => ({
      ...prev,
      [targetName]: selections
    }));
  };

  const handleDeleteMapping = (targetName) => {
    setColumnMappings(prev => {
      const updated = { ...prev };
      delete updated[targetName];
      return updated;
    });
  };

  // Action: Modify configs
  const handleRowConfigChange = (fileId, newConfig) => {
    setRowConfigs(prev => ({
      ...prev,
      [fileId]: newConfig
    }));
  };

  // Action: Change sheet location assignment
  const handleLocationChange = (fileId, sheetName, locationValue) => {
    setSheetLocations(prev => ({
      ...prev,
      [fileId]: {
        ...(prev[fileId] || {}),
        [sheetName]: locationValue
      }
    }));
  };

  // Action: Export merged output to download
  const handleExport = () => {
    if (mergedData.length === 0) {
      triggerNotification('error', 'Cannot export: Merged dataset contains no rows.');
      return;
    }

    if (!rangeValidation.isValid) {
      triggerNotification('error', rangeValidation.errorMsg);
      return;
    }

    setIsExporting(true);
    setTimeout(() => {
      try {
        exportMergedToExcel(mergedData, selectedColumns, outputFileName);
        triggerNotification('success', `Merged file "${outputFileName}.xlsx" downloaded successfully!`);
      } catch (err) {
        triggerNotification('error', `Export failed: ${err.message}`);
      } finally {
        setIsExporting(false);
      }
    }, 300);
  };

  // Action: Reset/Clear Workspace (Wipes database and React states)
  const handleReset = async () => {
    try {
      await clearWorkspace();
      setFiles([]);
      setRowConfigs({});
      setColumnMappings({});
      setSheetLocations({});
      setOutputFileName('merged_excel_output');
      triggerNotification('info', 'Workspace database cleared successfully.');
    } catch (err) {
      triggerNotification('error', `Reset failed: ${err.message}`);
    }
  };

  return (
    <div className="app-shell">
      {/* Floating Toast Notification messages */}
      {notification && (
        <div className={`toast-notification toast-${notification.type}`} role="alert">
          {notification.type === 'success' ? (
            <Check size={18} className="toast-icon" />
          ) : (
            <AlertCircle size={18} className="toast-icon" />
          )}
          <span className="toast-text">{notification.message}</span>
        </div>
      )}

      {/* Hero Branding Dashboard Banner */}
      <header className="dashboard-banner">
        <div className="header-branding">
          <div className="branding-icon-glow">
            <GitMerge size={32} className="branding-icon" />
          </div>
          <div>
            <h1 className="dashboard-title">Excel Extractor & Merger</h1>
            <p className="dashboard-subtitle">Combine multiple workbooks vertically, align dynamic schema, and filter precise rows instantly.</p>
          </div>
        </div>

        <div className="header-meta">
          <button 
            type="button" 
            className="btn btn-secondary btn-danger-hover" 
            onClick={handleReset}
            disabled={files.length === 0 && Object.keys(columnMappings).length === 0}
            id="reset-all-button"
            title="Wipe browser database and reset workspace"
          >
            <RefreshCw size={14} className="btn-icon-inline" />
            Clear Workspace
          </button>
        </div>
      </header>

      {/* Main Grid Workspace */}
      <main className="dashboard-workspace">
        <div className="grid-left-col">
          {/* File Picker */}
          <FilePicker 
            files={files}
            rowConfigs={rowConfigs}
            sheetLocations={sheetLocations}
            uniqueLocations={uniqueLocations}
            onLocationChange={handleLocationChange}
            onConfigChange={handleRowConfigChange}
            onFilesAdded={handleFilesAdded}
            onRemoveFile={handleRemoveFile}
            onSheetChanged={handleSheetChanged}
            onError={(msg) => triggerNotification('error', msg)}
            onNotification={triggerNotification}
          />

          {/* Row Ranges Configurations */}
          {files.length > 0 && (
            <RowRangeSelector 
              files={files}
              rowConfigs={rowConfigs}
              onConfigChange={handleRowConfigChange}
            />
          )}
        </div>

        <div className="grid-right-col">
          {/* Column Mapping Studio */}
          {files.length > 0 ? (
            <ColumnMapper 
              files={files}
              rowConfigs={rowConfigs}
              columnMappings={columnMappings}
              onMappingAdd={handleAddMapping}
              onMappingDelete={handleDeleteMapping}
              onNotification={triggerNotification}
            />
          ) : (
            <div className="workspace-empty-card card">
              <Sparkles size={48} className="empty-sparkle-icon" />
              <h4>Ready to Merge Sheets?</h4>
              {isInitialLoad ? (
                <p>Loading saved workspace database, please wait...</p>
              ) : (
                <p>Upload your Excel workbooks on the left to automatically parse active tabs, suggest custom header offsets, and configure unified column mappings.</p>
              )}
            </div>
          )}

          {/* Export Settings Card */}
          {files.length > 0 && (
            <div className="export-settings-card card">
              <h3 className="section-title">
                <Settings size={18} className="title-icon" />
                Export Settings
              </h3>
              
              <div className="export-settings-form">
                <div className="export-input-wrapper">
                  <label className="export-label" htmlFor="filename-input">Output Filename</label>
                  <div className="filename-input-container">
                    <input 
                      type="text"
                      className="filename-textbox"
                      id="filename-input"
                      value={outputFileName}
                      onChange={(e) => setOutputFileName(e.target.value)}
                      placeholder="merged_excel_output"
                    />
                    <span className="file-extension-badge">.xlsx</span>
                  </div>
                </div>

                <button 
                  type="button"
                  className="btn btn-primary btn-large btn-glow"
                  onClick={handleExport}
                  disabled={isExporting || mergedData.length === 0 || !rangeValidation.isValid}
                  id="export-trigger-button"
                >
                  {isExporting ? (
                    <>
                      <RefreshCw size={18} className="btn-icon-inline animate-spin" />
                      Exporting...
                    </>
                  ) : (
                    <>
                      <FileDown size={18} className="btn-icon-inline" />
                      Merge & Export Excel
                    </>
                  )}
                </button>
              </div>

              {!rangeValidation.isValid && (
                <div className="validation-error-banner">
                  <AlertCircle size={14} className="inline-icon" />
                  <span><strong>Range Error:</strong> {rangeValidation.errorMsg}</span>
                </div>
              )}
            </div>
          )}
        </div>
      </main>

      {/* Dynamic Live Preview Section */}
      <section className="dashboard-preview-container">
        <PreviewTable 
          mergedData={mergedData}
          selectedColumns={selectedColumns}
        />
      </section>

      {/* Footer bar */}
      <footer className="workspace-footer">
        <span>Excel Processor Utility v1.4.0</span>
        <span>Persistent Client-Side File Welding Pipeline with Unified Mapping Studio</span>
      </footer>
    </div>
  );
}
