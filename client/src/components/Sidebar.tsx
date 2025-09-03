import React, { useState, useRef } from 'react';
import { UserSessionState } from '../types';

interface SidebarProps {
  isCollapsed: boolean;
  onToggle: () => void;
  state: UserSessionState;
  updateState: (newState: Partial<UserSessionState>) => void;
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  isCollapsed,
  onToggle,
  state,
  updateState,
  showNotification
}) => {
  const [dragOver, setDragOver] = useState(false);
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileConfigChange = (field: 'tableName' | 'sheetName', value: string) => {
    updateState({
      fileConfigs: {
        ...state.fileConfigs,
        [field]: value
      }
    });
  };

  const handleDocPropChange = (field: keyof typeof state.docProps, value: string) => {
    updateState({
      docProps: {
        ...state.docProps,
        [field]: value
      }
    });
  };

  const preventDefaults = (e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  };

  const handleDragEnter = (e: React.DragEvent) => {
    preventDefaults(e);
    setDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    preventDefaults(e);
    setDragOver(false);
  };

  const handleDragOver = (e: React.DragEvent) => {
    preventDefaults(e);
    setDragOver(true);
  };

  const handleDrop = (e: React.DragEvent) => {
    preventDefaults(e);
    setDragOver(false);
    
    const files = Array.from(e.dataTransfer.files);
    if (files.length > 0) {
      handleFiles(files);
    }
  };

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      handleFiles(Array.from(e.target.files));
    }
  };

  const handleFiles = (files: File[]) => {
    files.forEach(handleFile);
  };

  const handleFile = (file: File) => {
    // Check if file is Excel format
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
      'application/vnd.ms-excel' // .xls
    ];
    
    if (!validTypes.includes(file.type) && !file.name.match(/\.(xlsx|xls)$/i)) {
      showNotification('Please select only Excel files (.xlsx or .xls)', 'error');
      return;
    }
    
    // Show file upload success
    showNotification(`File "${file.name}" uploaded successfully`, 'success');
    setUploadedFile(file);
    
    // Add file to session state uploads array
    updateState({
      uploads: [...state.uploads, file]
    });
    
    console.log('File uploaded:', file.name, 'Size:', file.size, 'Type:', file.type);
  };

  const handleDragDropClick = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className={`right-sidebar ${isCollapsed ? 'collapsed' : ''}`}>
      <div className="sidebar-toggle" onClick={onToggle}>
        <span className="toggle-icon">◀</span>
      </div>
      <div className="sidebar-content">
        <div className="sidebar-section">
          <h3>File Configs</h3>
        
          {/* XLSX Upload Section */}
          <div className="upload-section">
            <div 
              className={`drag-drop-area ${dragOver ? 'dragover' : ''}`}
              onDragEnter={handleDragEnter}
              onDragLeave={handleDragLeave}
              onDragOver={handleDragOver}
              onDrop={handleDrop}
              onClick={handleDragDropClick}
            >
              <div className="drag-drop-content">
                {uploadedFile ? (
                  <>
                    <span className="upload-icon">✅</span>
                    <p><strong>{uploadedFile.name}</strong></p>
                    <p className="upload-hint">File uploaded successfully</p>
                    <p className="upload-hint">Click to upload another file</p>
                  </>
                ) : (
                  <>
                    <span className="upload-icon">📊</span>
                    <p>Drag & drop XLSX files here</p>
                    <p className="upload-hint">or click to browse</p>
                  </>
                )}
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls"
                multiple
                style={{ display: 'none' }}
                onChange={handleFileInputChange}
              />
            </div>
            
            <div className="config-inputs">
              <div className="input-group">
                <label htmlFor="tableName">Table Name:</label>
                <input
                  type="text"
                  id="tableName"
                  value={state.fileConfigs.tableName}
                  className="config-input"
                  onChange={(e) => handleFileConfigChange('tableName', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="sheetName">Sheet Name:</label>
                <input
                  type="text"
                  id="sheetName"
                  value={state.fileConfigs.sheetName}
                  className="config-input"
                  onChange={(e) => handleFileConfigChange('sheetName', e.target.value)}
                />
              </div>
            </div>
          </div>
          
          {/* Doc Props Section */}
          <div className="doc-props-section">
            <h4>Doc Props</h4>
            
            <div className="prop-inputs">
              <div className="input-group">
                <label htmlFor="docTitle">Title:</label>
                <input
                  type="text"
                  id="docTitle"
                  className="prop-input"
                  placeholder="Document title"
                  value={state.docProps.title}
                  onChange={(e) => handleDocPropChange('title', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docSubject">Subject:</label>
                <input
                  type="text"
                  id="docSubject"
                  className="prop-input"
                  placeholder="Document subject"
                  value={state.docProps.subject}
                  onChange={(e) => handleDocPropChange('subject', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docKeywords">Keywords:</label>
                <input
                  type="text"
                  id="docKeywords"
                  className="prop-input"
                  placeholder="Search keywords"
                  value={state.docProps.keywords}
                  onChange={(e) => handleDocPropChange('keywords', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docCreatedBy">Created By:</label>
                <input
                  type="text"
                  id="docCreatedBy"
                  className="prop-input"
                  placeholder="Author name"
                  value={state.docProps.createdBy}
                  onChange={(e) => handleDocPropChange('createdBy', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docDescription">Description:</label>
                <textarea
                  id="docDescription"
                  className="prop-input"
                  placeholder="Document description"
                  rows={3}
                  value={state.docProps.description}
                  onChange={(e) => handleDocPropChange('description', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docLastModifiedBy">Last Modified By:</label>
                <input
                  type="text"
                  id="docLastModifiedBy"
                  className="prop-input"
                  placeholder="Last editor"
                  value={state.docProps.lastModifiedBy}
                  onChange={(e) => handleDocPropChange('lastModifiedBy', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docCategory">Category:</label>
                <input
                  type="text"
                  id="docCategory"
                  className="prop-input"
                  placeholder="Document category"
                  value={state.docProps.category}
                  onChange={(e) => handleDocPropChange('category', e.target.value)}
                />
              </div>
              
              <div className="input-group">
                <label htmlFor="docRevision">Revision:</label>
                <input
                  type="text"
                  id="docRevision"
                  className="prop-input"
                  placeholder="Version number"
                  value={state.docProps.revision}
                  onChange={(e) => handleDocPropChange('revision', e.target.value)}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Sidebar;