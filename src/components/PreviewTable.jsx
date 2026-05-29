import React, { useState, useMemo } from 'react';
import { Eye, ArrowUpDown, ChevronLeft, ChevronRight, ChevronsLeft, ChevronsRight, EyeOff, LayoutGrid } from 'lucide-react';

export default function PreviewTable({ mergedData, selectedColumns }) {
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const [sortConfig, setSortConfig] = useState({ key: null, direction: null });

  // Handle Header Click for Sorting
  const handleSort = (colKey) => {
    let direction = 'asc';
    if (sortConfig.key === colKey) {
      if (sortConfig.direction === 'asc') {
        direction = 'desc';
      } else if (sortConfig.direction === 'desc') {
        direction = null; // Reset sort
      }
    }
    setSortConfig({ key: direction ? colKey : null, direction });
    setCurrentPage(1); // Reset page on sort
  };

  // Sort the merged data
  const sortedData = useMemo(() => {
    if (!sortConfig.key || !sortConfig.direction) return mergedData;

    const key = sortConfig.key;
    const direction = sortConfig.direction === 'asc' ? 1 : -1;

    return [...mergedData].sort((a, b) => {
      const valA = a[key];
      const valB = b[key];

      // Float empty values to bottom
      if (valA === "" || valA === undefined || valA === null) return 1;
      if (valB === "" || valB === undefined || valB === null) return -1;

      // Type-specific comparison
      if (valA instanceof Date && valB instanceof Date) {
        return (valA.getTime() - valB.getTime()) * direction;
      }
      
      const numA = Number(valA);
      const numB = Number(valB);
      if (!isNaN(numA) && !isNaN(numB)) {
        return (numA - numB) * direction;
      }

      return String(valA).localeCompare(String(valB), undefined, { numeric: true, sensitivity: 'base' }) * direction;
    });
  }, [mergedData, sortConfig]);

  // Pagination bounds
  const totalRows = sortedData.length;
  const totalPages = Math.ceil(totalRows / pageSize) || 1;

  const paginatedData = useMemo(() => {
    const startIdx = (currentPage - 1) * pageSize;
    return sortedData.slice(startIdx, startIdx + pageSize);
  }, [sortedData, currentPage, pageSize]);

  // Reset page size
  const handlePageSizeChange = (e) => {
    const newSize = parseInt(e.target.value, 10);
    setPageSize(newSize);
    setCurrentPage(1);
  };

  // Placeholder if no mapping rules defined yet
  if (selectedColumns.length === 0) {
    return (
      <div className="preview-emptycard card">
        <h3 className="section-title">
          <Eye size={18} className="title-icon" />
          Live Dataset Preview
        </h3>
        <div className="table-empty-state">
          <LayoutGrid size={48} className="empty-state-icon" style={{ color: 'rgba(139, 92, 246, 0.15)' }} />
          <p className="empty-state-text">No column mappings configured yet</p>
          <p className="empty-state-subtext">
            Add unified column mapping rules in the <strong>Column Mapping Studio</strong> above. 
            Only mapped target columns will be merged and exported.
          </p>
        </div>
      </div>
    );
  }

  // Placeholder if mapping rules exist but range or parsing results are empty
  if (mergedData.length === 0) {
    return (
      <div className="preview-emptycard card">
        <h3 className="section-title">
          <Eye size={18} className="title-icon" />
          Live Dataset Preview
        </h3>
        <div className="table-empty-state">
          <EyeOff size={48} className="empty-state-icon" />
          <p className="empty-state-text">No merged records to display</p>
          <p className="empty-state-subtext">
            Verify that your upload is active, and range coordinates do not contain validation errors.
          </p>
        </div>
      </div>
    );
  }

  return (
    <div className="preview-section card">
      <div className="preview-section-header">
        <h3 className="section-title">
          <Eye size={18} className="title-icon" />
          Live Dataset Preview
        </h3>
      </div>

      <div className="preview-meta-stats">
        <span>Total Rows Combined: <strong className="stat-count text-primary">{totalRows}</strong></span>
        <span>Mapped Output Columns: <strong className="stat-count text-accent">{selectedColumns.length}</strong></span>
      </div>

      {/* Dynamic Scrollable Table Container */}
      <div className="table-scroll-container">
        <table className="preview-table">
          <thead>
            <tr>
              {/* Audit Metadata Columns */}
              <th className="audit-col" title="Origin filename">Source File</th>
              <th className="audit-col" title="Worksheet name">Source Sheet</th>
              <th className="audit-col text-center" title="Original spreadsheet index">Original Row</th>
              
              {/* Target Columns */}
              {selectedColumns.map(col => {
                const isSorted = sortConfig.key === col;
                
                return (
                  <th 
                    key={col} 
                    onClick={() => handleSort(col)}
                    className={`sortable-header ${isSorted ? 'sorted-active' : ''}`}
                    title="Click to sort column"
                  >
                    <div className="header-cell-content">
                      <span className="header-text" title={col}>{col}</span>
                      <ArrowUpDown size={12} className={`sort-arrow-icon ${isSorted ? 'active' : ''}`} />
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody>
            {paginatedData.map((row, rowIdx) => (
              <tr key={rowIdx} className="table-data-row">
                {/* Audit values */}
                <td className="audit-col text-ellipsis" title={row._sourceFile}>
                  {row._sourceFile}
                </td>
                <td className="audit-col text-ellipsis" title={row._sourceSheet}>
                  {row._sourceSheet}
                </td>
                <td className="audit-col text-center font-mono">
                  {row._sourceExcelRow}
                </td>

                {/* Spreadsheet cell values */}
                {selectedColumns.map(col => {
                  const cellVal = row[col];
                  
                  // Safe string representation
                  let displayVal = "";
                  if (cellVal instanceof Date) {
                    displayVal = cellVal.toLocaleDateString();
                  } else if (cellVal !== undefined && cellVal !== null) {
                    displayVal = String(cellVal);
                  }

                  return (
                    <td 
                      key={col} 
                      className={displayVal === "" ? 'empty-cell' : ''}
                      title={displayVal || "(empty)"}
                    >
                      {displayVal || <span className="empty-placeholder">-</span>}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Pagination Controls */}
      <div className="pagination-bar">
        <div className="page-size-selector">
          <span>Rows per page:</span>
          <select 
            value={pageSize}
            onChange={handlePageSizeChange}
            className="page-dropdown"
          >
            <option value={5}>5</option>
            <option value={10}>10</option>
            <option value={25}>25</option>
            <option value={50}>50</option>
            <option value={100}>100</option>
          </select>
        </div>

        <div className="page-pagination-controls">
          <button 
            type="button" 
            className="btn-pagination" 
            onClick={() => setCurrentPage(1)} 
            disabled={currentPage === 1}
            title="First page"
          >
            <ChevronsLeft size={16} />
          </button>
          <button 
            type="button" 
            className="btn-pagination" 
            onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))} 
            disabled={currentPage === 1}
            title="Previous page"
          >
            <ChevronLeft size={16} />
          </button>
          
          <span className="page-indicator">
            Page <strong>{currentPage}</strong> of <strong>{totalPages}</strong>
          </span>

          <button 
            type="button" 
            className="btn-pagination" 
            onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))} 
            disabled={currentPage === totalPages}
            title="Next page"
          >
            <ChevronRight size={16} />
          </button>
          <button 
            type="button" 
            className="btn-pagination" 
            onClick={() => setCurrentPage(totalPages)} 
            disabled={currentPage === totalPages}
            title="Last page"
          >
            <ChevronsRight size={16} />
          </button>
        </div>

        <div className="pagination-stats">
          Showing {Math.min(totalRows, (currentPage - 1) * pageSize + 1)} - {Math.min(totalRows, currentPage * pageSize)} of {totalRows} rows
        </div>
      </div>
    </div>
  );
}
