import React, { useState, useMemo } from 'react';
import { 
  GitCommit, 
  Plus, 
  Trash2, 
  Network, 
  ArrowRightLeft, 
  FileSpreadsheet, 
  Link2,
  AlertCircle
} from 'lucide-react';
import { extractHeadersFromRow } from '../utils/excelProcessor';

export default function ColumnMapper({ 
  files, 
  rowConfigs, 
  columnMappings, 
  onMappingAdd, 
  onMappingDelete, 
  onNotification 
}) {
  const [targetName, setTargetName] = useState('');
  const [selections, setSelections] = useState({});
  const [isFormOpen, setIsFormOpen] = useState(false);

  // Compute available headers per file dynamically
  const fileHeadersMap = useMemo(() => {
    const map = {};
    files.forEach(file => {
      const activeSheetData = file.sheets[file.activeSheet];
      if (activeSheetData && activeSheetData.rawGrid) {
        const config = rowConfigs[file.id] || { headerRow: 1 };
        const headerRowIdx = Math.max(0, config.headerRow - 1);
        map[file.id] = extractHeadersFromRow(activeSheetData.rawGrid, headerRowIdx);
      } else {
        map[file.id] = [];
      }
    });
    return map;
  }, [files, rowConfigs]);

  // Handle dropdown mapping change
  const handleSelectChange = (fileId, colName) => {
    setSelections(prev => ({
      ...prev,
      [fileId]: colName || ""
    }));
  };

  // Trigger Unified Column creation
  const handleCreateMapping = (e) => {
    e.preventDefault();
    const trimmedTargetName = targetName.trim();

    if (!trimmedTargetName) {
      onNotification('error', 'Please enter a name for the unified column.');
      return;
    }

    // Check if target name conflicts with existing defined aliases
    if (columnMappings[trimmedTargetName]) {
      onNotification('error', `A unified column named "${trimmedTargetName}" already exists.`);
      return;
    }

    // Filter out empty selections
    const activeSelections = {};
    Object.keys(selections).forEach(fId => {
      if (selections[fId]) {
        activeSelections[fId] = selections[fId];
      }
    });

    if (Object.keys(activeSelections).length === 0) {
      onNotification('error', 'Please select at least one source column mapping.');
      return;
    }

    // Callback to save
    onMappingAdd(trimmedTargetName, activeSelections);

    // Reset forms state
    setTargetName('');
    setSelections({});
    setIsFormOpen(false);
    onNotification('success', `Unified column "${trimmedTargetName}" created successfully.`);
  };

  const hasMappings = Object.keys(columnMappings).length > 0;

  return (
    <div className="column-mapper-card card">
      <h3 className="section-title">
        <Network size={18} className="title-icon" />
        Column Mapping Studio
      </h3>
      
      <p className="range-section-helper">
        Combine differently named columns from separate files (e.g. <em>"actual wages"</em> and <em>"Total Gross"</em>) into a single, consolidated output column during vertical merging.
      </p>

      {/* Button to open inline form creator */}
      {!isFormOpen && (
        <button 
          type="button" 
          className="btn btn-secondary btn-small"
          onClick={() => setIsFormOpen(true)}
          disabled={files.length < 2}
          title="Group multiple headers across sheets"
        >
          <Plus size={14} className="btn-icon-inline" />
          Create Unified Column Alias
        </button>
      )}

      {/* Inline Creation Form Workspace */}
      {isFormOpen && (
        <form onSubmit={handleCreateMapping} className="alias-creator-form">
          <div className="alias-form-header">
            <span>Define Column Aliases</span>
            <button 
              type="button" 
              className="btn-close-form" 
              onClick={() => setIsFormOpen(false)}
            >
              Cancel
            </button>
          </div>

          <div className="alias-form-body">
            {/* Unified target name input */}
            <div className="input-group">
              <label className="export-label" htmlFor="alias-target-input">
                Unified Output Column Name
              </label>
              <input 
                type="text"
                id="alias-target-input"
                className="filename-textbox textbox-bordered"
                value={targetName}
                onChange={(e) => setTargetName(e.target.value)}
                placeholder="e.g. Gross Salary"
                required
              />
            </div>

            {/* Selection mapping list per uploaded file */}
            <div className="file-selections-list">
              <span className="selections-title">Select source columns per file:</span>
              
              {files.map(file => {
                const headers = fileHeadersMap[file.id] || [];
                return (
                  <div key={file.id} className="file-dropdown-row">
                    <div className="dropdown-file-meta">
                      <FileSpreadsheet size={14} className="excel-file-icon" />
                      <span className="dropdown-filename" title={file.name}>
                        {file.name}
                      </span>
                    </div>

                    <select 
                      value={selections[file.id] || ""}
                      onChange={(e) => handleSelectChange(file.id, e.target.value)}
                      className="sheet-dropdown alias-dropdown"
                    >
                      <option value="">(Do not map)</option>
                      {headers.map(h => (
                        <option key={h.key} value={h.key}>
                          {h.key} (Col {h.letter})
                        </option>
                      ))}
                    </select>
                  </div>
                );
              })}
            </div>

            <button 
              type="submit" 
              className="btn btn-primary btn-small btn-glow"
              style={{ marginTop: '0.5rem', width: '100%' }}
            >
              <Link2 size={14} className="btn-icon-inline" />
              Save Mapping Rules
            </button>
          </div>
        </form>
      )}

      {/* Active Mappings Visualizer Board */}
      {hasMappings && (
        <div className="active-mappings-panel">
          <span className="file-list-header">Active Mappings Rules</span>
          
          <div className="mappings-flow-grid">
            {Object.keys(columnMappings).map(target => {
              const fileMap = columnMappings[target];
              return (
                <div key={target} className="mapping-flow-card">
                  {/* Title Bar containing clean target name */}
                  <div className="flow-card-header">
                    <div className="flow-target-branding">
                      <ArrowRightLeft size={14} className="text-accent" />
                      <span className="flow-target-name">{target}</span>
                      <span className="flow-target-badge">Unified Target</span>
                    </div>

                    <button 
                      type="button"
                      onClick={() => onMappingDelete(target)}
                      className="btn-icon btn-danger"
                      title="Delete Mapping Rule"
                      aria-label={`Delete mapping rule for ${target}`}
                    >
                      <Trash2 size={14} />
                    </button>
                  </div>

                  {/* Connecting visual flows lists */}
                  <div className="flow-connections-list">
                    {Object.keys(fileMap).map(fId => {
                      const file = files.find(f => f.id === fId);
                      if (!file) return null;
                      return (
                        <div key={fId} className="flow-connection-item">
                          <div className="flow-source-file text-ellipsis">
                            <span className="flow-bullet">•</span>
                            {file.name}
                          </div>
                          
                          <div className="flow-connector-line">
                            <span className="connector-arrow">➡️</span>
                          </div>

                          <div className="flow-source-column text-ellipsis">
                            {fileMap[fId]}
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}
