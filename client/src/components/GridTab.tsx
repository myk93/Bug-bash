import React, { useState, useEffect, useCallback, useRef } from 'react';
import { UserSessionState } from '../types';

interface GridTabProps {
  state: UserSessionState;
  updateState: (newState: Partial<UserSessionState>) => void;
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
}

export const GridTab: React.FC<GridTabProps> = ({ state, updateState, showNotification }) => {
  const [gridData, setGridData] = useState<string[][]>([]);
  const [currentRows, setCurrentRows] = useState(3);
  const [currentCols, setCurrentCols] = useState(3);
  const [rowCount, setRowCount] = useState(3);
  const [colCount, setColCount] = useState(3);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Initialize grid with default 3x3
  useEffect(() => {
    if (state.gridData) {
      setGridData(state.gridData.data);
      setCurrentRows(state.gridData.rows);
      setCurrentCols(state.gridData.cols);
      setRowCount(state.gridData.rows);
      setColCount(state.gridData.cols);
    } else {
      createGrid(3, 3);
    }
  }, []);

  const createGrid = useCallback((rows: number, cols: number) => {
    console.log('Creating grid:', rows, 'x', cols);
    
    // Initialize or adjust grid data
    const newGridData: string[][] = [];
    for (let i = 0; i < rows; i++) {
      newGridData[i] = [];
      for (let j = 0; j < cols; j++) {
        if (gridData[i] && gridData[i][j] !== undefined) {
          newGridData[i][j] = gridData[i][j];
        } else {
          newGridData[i][j] = '';
        }
      }
    }
    
    setGridData(newGridData);
    setCurrentRows(rows);
    setCurrentCols(cols);
    
    // Update session state
    updateState({
      gridData: {
        data: newGridData,
        rows: rows,
        cols: cols
      }
    });
  }, [gridData, updateState]);

  const updateGridInSession = useCallback((newData: string[][]) => {
    updateState({
      gridData: {
        data: newData,
        rows: currentRows,
        cols: currentCols
      }
    });
  }, [updateState, currentRows, currentCols]);

  const handleCellChange = (row: number, col: number, value: string) => {
    const newGridData = [...gridData];
    newGridData[row][col] = value;
    setGridData(newGridData);
    updateGridInSession(newGridData);
  };

  const handleUpdateGrid = () => {
    if (rowCount < 1 || rowCount > 20 || colCount < 1 || colCount > 20) {
      showNotification('Grid size must be between 1x1 and 20x20', 'error');
      return;
    }
    
    createGrid(rowCount, colCount);
    showNotification(`Grid updated to ${rowCount}×${colCount}`, 'success');
  };

  const handleClearGrid = () => {
    const confirmed = window.confirm('Clear all grid data? This action cannot be undone.');
    if (!confirmed) return;
    
    const clearedData = Array(currentRows).fill(null).map(() => Array(currentCols).fill(''));
    setGridData(clearedData);
    updateGridInSession(clearedData);
    showNotification('Grid cleared successfully', 'success');
  };

  const handleImportGrid = () => {
    fileInputRef.current?.click();
  };

  const handleFileImport = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const content = e.target?.result as string;
        const rows = content.split('\n').filter(row => row.trim());
        const importedData = rows.map(row => {
          // Simple CSV parsing
          const cells: string[] = [];
          let current = '';
          let inQuotes = false;
          
          for (let i = 0; i < row.length; i++) {
            const char = row[i];
            if (char === '"' && (i === 0 || row[i-1] === ',')) {
              inQuotes = true;
            } else if (char === '"' && inQuotes && (i === row.length - 1 || row[i+1] === ',')) {
              inQuotes = false;
            } else if (char === ',' && !inQuotes) {
              cells.push(current.replace(/""/g, '"'));
              current = '';
            } else if (!(char === '"' && (i === 0 || row[i-1] === ',' || i === row.length - 1 || row[i+1] === ','))) {
              current += char;
            }
          }
          cells.push(current.replace(/""/g, '"'));
          return cells;
        });
        
        if (importedData.length > 0) {
          const maxCols = Math.max(...importedData.map(row => row.length));
          const newRows = Math.min(importedData.length, 20);
          const newCols = Math.min(maxCols, 20);
          
          setRowCount(newRows);
          setColCount(newCols);
          
          // Create new grid with imported data
          const newGridData: string[][] = [];
          for (let i = 0; i < newRows; i++) {
            newGridData[i] = [];
            for (let j = 0; j < newCols; j++) {
              if (importedData[i] && importedData[i][j] !== undefined) {
                newGridData[i][j] = importedData[i][j];
              } else {
                newGridData[i][j] = '';
              }
            }
          }
          
          setGridData(newGridData);
          setCurrentRows(newRows);
          setCurrentCols(newCols);
          updateGridInSession(newGridData);
          
          showNotification('Grid data imported successfully', 'success');
        }
      } catch (error) {
        console.error('Import error:', error);
        showNotification('Error importing file. Please check file format.', 'error');
      }
    };
    reader.readAsText(file);
    
    // Clear the input
    event.target.value = '';
  };

  const handleGridViewToggle = (checked: boolean) => {
    updateState({
      gridView: {
        ...(state.gridView || { isGridView: false, promoteHeaders: false, adjustColumnNames: false }),
        isGridView: checked
      }
    });
    
    if (checked) {
      showNotification('Switched to Grid view', 'info');
    } else {
      showNotification('Switched to HTML Table view', 'info');
    }
  };

  const handlePromoteHeaders = (checked: boolean) => {
    updateState({
      gridView: {
        ...(state.gridView || { isGridView: false, promoteHeaders: false, adjustColumnNames: false }),
        promoteHeaders: checked
      }
    });
    
    if (checked) {
      showNotification('Headers will be promoted from first row', 'info');
    }
  };

  const handleAdjustColumnNames = (checked: boolean) => {
    updateState({
      gridView: {
        ...(state.gridView || { isGridView: false, promoteHeaders: false, adjustColumnNames: false }),
        adjustColumnNames: checked
      }
    });
    
    if (checked) {
      showNotification('Column names will be adjusted for duplicates/invalid names', 'info');
    }
  };

  return (
    <div id="Grid" className="tabcontent active">
      <h3>From Grid</h3>
      
      {/* Grid View Toggle and Options */}
      <div className="grid-view-container">
        <div className="grid-view-toggle">
          <input
            type="checkbox"
            id="gridViewToggle"
            className="grid-view-toggle-input"
            checked={state.gridView?.isGridView || false}
            onChange={(e) => handleGridViewToggle(e.target.checked)}
          />
          <label htmlFor="gridViewToggle" className="grid-view-toggle-label">
            <span className="grid-view-toggle-text left">HTML Table</span>
            <span className="grid-view-toggle-slider"></span>
            <span className="grid-view-toggle-text right">As Grid</span>
          </label>
        </div>
        
        {/* Grid Options (shown only when "As Grid" is selected) */}
        <div
          className="grid-options"
          id="gridOptions"
          style={{ display: (state.gridView?.isGridView || false) ? 'flex' : 'none' }}
        >
          <div className="grid-option-item">
            <label className="grid-option-label" title="Use first row as headers">
              <input
                type="checkbox"
                id="promoteHeaders"
                className="grid-option-checkbox"
                checked={state.gridView?.promoteHeaders || false}
                onChange={(e) => handlePromoteHeaders(e.target.checked)}
              />
              <span className="checkbox-custom"></span>
              <span className="option-text">Promote Headers</span>
            </label>
          </div>
          <div className="grid-option-item">
            <label className="grid-option-label" title="Fix duplicate/invalid names">
              <input
                type="checkbox"
                id="adjustColumnNames"
                className="grid-option-checkbox"
                checked={state.gridView?.adjustColumnNames || false}
                onChange={(e) => handleAdjustColumnNames(e.target.checked)}
              />
              <span className="checkbox-custom"></span>
              <span className="option-text">Adjust Column Names</span>
            </label>
          </div>
        </div>
      </div>
      
      {/* Grid Controls */}
      <div className="grid-controls">
        <div className="control-group">
          <label htmlFor="rowCount">Rows:</label>
          <input
            type="number"
            id="rowCount"
            min="1"
            max="20"
            value={rowCount}
            onChange={(e) => setRowCount(parseInt(e.target.value) || 1)}
            className="control-input"
          />
        </div>
        <div className="control-group">
          <label htmlFor="colCount">Columns:</label>
          <input
            type="number"
            id="colCount"
            min="1"
            max="20"
            value={colCount}
            onChange={(e) => setColCount(parseInt(e.target.value) || 1)}
            className="control-input"
          />
        </div>
        <button onClick={handleUpdateGrid} className="update-btn">Update Grid</button>
        <button onClick={handleClearGrid} className="clear-btn">Clear All</button>
      </div>
      
      {/* Editable Grid Table */}
      <div className="grid-table-container">
        <table className="editable-grid">
          <tbody>
            {gridData.map((row, rowIndex) => (
              <tr key={rowIndex}>
                {row.map((cell, colIndex) => (
                  <td key={colIndex}>
                    <textarea
                      className="grid-cell"
                      value={cell}
                      onChange={(e) => handleCellChange(rowIndex, colIndex, e.target.value)}
                      placeholder={`R${rowIndex + 1}C${colIndex + 1}`}
                      rows={1}
                      onFocus={(e) => e.target.select()}
                    />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {/* Grid Actions */}
      <div className="grid-actions">
        <button onClick={handleImportGrid} className="action-btn">Import Data</button>
        <span className="grid-info">{currentRows} rows × {currentCols} columns</span>
      </div>
      
      {/* Hidden file input for import */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".csv,.txt"
        style={{ display: 'none' }}
        onChange={handleFileImport}
      />
    </div>
  );
};

export default GridTab;