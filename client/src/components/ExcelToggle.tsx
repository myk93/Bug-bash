import React, { useState } from 'react';
import { UserSessionState } from '../types';
import { generateQueryWorkbook, handleExcelToggle } from '../utils/exportUtils';

interface ExcelToggleProps {
  excelToggle: boolean;
  onToggleChange: (checked: boolean) => void;
  onExport: () => void;
  onReset: () => void;
  state: UserSessionState;
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
}

export const ExcelToggle: React.FC<ExcelToggleProps> = ({
  excelToggle,
  onToggleChange,
  onExport,
  onReset,
  state,
  showNotification
}) => {
  const [isExporting, setIsExporting] = useState(false);
  const [isResetting, setIsResetting] = useState(false);

  const handleExportClick = async () => {
    try {
      setIsExporting(true);
      console.log('Exporting data...');
      showNotification('Preparing export...', 'info');
      
      // Get the current active tab from session manager
      const activeTab = state.activeTab;
      console.log('Active tab for export:', activeTab);
      
      let blob: Blob | null = null;
      
      console.log('Generating workbook for active tab:', activeTab);
      if (activeTab === 'pq-query') {
        // For PQQuery tab: call generateQueryWorkbook to get the blob
        console.log('Generating workbook for PQQuery tab...');
        blob = await generateQueryWorkbook(state);
      } else if (activeTab === 'grid') {
        // For Grid tab: show alert that Grid export is not implemented yet
        showNotification('Grid export is not implemented yet', 'warning');
        console.log('Grid export not yet implemented');
        return;
      } else {
        // For other tabs
        showNotification(`Export not available for ${activeTab} tab`, 'warning');
        console.log(`Export not available for tab: ${activeTab}`);
        return;
      }
      
      if (blob) {
        // Call handleExcelToggle with the generated blob
        console.log('Calling handleExcelToggle with generated blob...');
        await handleExcelToggle(blob, excelToggle, activeTab);
        
        const message = excelToggle 
          ? `Workbook downloaded successfully`
          : `Opened in Excel Web successfully`;
        showNotification(message, 'success');
      } else {
        throw new Error('Failed to generate workbook blob');
      }
      
    } catch (error) {
      console.error('Error exporting data:', error);
      const message = error instanceof Error ? error.message : 'Unknown error';
      showNotification(`Export failed: ${message}`, 'error');
    } finally {
      setIsExporting(false);
    }
  };

  const handleResetClick = async () => {
    // Show confirmation dialog
    const confirmed = window.confirm(
      'Reset your workspace? This will clear all your data.\n\n' +
      'This action cannot be undone. Your session will be reset to its default state.'
    );
    
    if (!confirmed) {
      return;
    }
    
    try {
      setIsResetting(true);
      showNotification('Resetting workspace...', 'info');
      
      // Call the parent reset function
      await onReset();
      
      showNotification('Workspace reset successfully', 'success');
      
      // Reload the page to show clean state
      setTimeout(() => {
        window.location.reload();
      }, 1000);
    } catch (error) {
      console.error('Error during reset:', error);
      showNotification('An error occurred while resetting. Please try again.', 'error');
    } finally {
      setIsResetting(false);
    }
  };

  return (
    <div className="excel-toggle-container">
      <div className="excel-toggle">
        <input
          type="checkbox"
          id="excelToggle"
          className="excel-toggle-input"
          checked={excelToggle}
          onChange={(e) => onToggleChange(e.target.checked)}
        />
        <label htmlFor="excelToggle" className="excel-toggle-label">
          <span className="excel-toggle-text left">Excel Web</span>
          <span className="excel-toggle-slider"></span>
          <span className="excel-toggle-text right">Download</span>
        </label>
        <button
          id="exportButton"
          className="export-button"
          type="button"
          onClick={handleExportClick}
          disabled={isExporting}
        >
          <span className="export-icon">📊</span>
          {isExporting ? 'Exporting...' : 'Export'}
        </button>
      </div>
      <button
        id="resetButton"
        className="reset-button"
        type="button"
        onClick={handleResetClick}
        disabled={isResetting}
      >
        <span className="reset-icon">{isResetting ? '⏳' : '⚠️'}</span>
        {isResetting ? 'Resetting...' : 'Reset'}
      </button>
    </div>
  );
};

export default ExcelToggle;