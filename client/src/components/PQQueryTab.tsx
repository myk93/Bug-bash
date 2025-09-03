import React from 'react';
import { UserSessionState } from '../types';

interface PQQueryTabProps {
  state: UserSessionState;
  updateState: (newState: Partial<UserSessionState>) => void;
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
}

const examples = {
  '1': '// Example 1: Basic table query\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]\nin\n    Source',
  '2': '// Example 2: Filter and transform\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    FilteredRows = Table.SelectRows(Source, each ([Column1] <> null))\nin\n    FilteredRows',
  '3': '// Example 3: Add custom column\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    AddedCustom = Table.AddColumn(Source, "Custom", each "Value")\nin\n    AddedCustom',
  '4': '// Example 4: Group and aggregate\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    GroupedRows = Table.Group(Source, {"Column1"}, {{"Count", Table.RowCount, Int64.Type}})\nin\n    GroupedRows'
};

export const PQQueryTab: React.FC<PQQueryTabProps> = ({ state, updateState, showNotification }) => {
  const handleQueryMashupChange = (value: string) => {
    updateState({
      pqQuery: {
        ...state.pqQuery,
        queryMashup: value
      }
    });
  };

  const handleRefreshOnOpenChange = (checked: boolean) => {
    updateState({
      pqQuery: {
        ...state.pqQuery,
        refreshOnOpen: checked
      }
    });
  };

  const handleQueryNameChange = (value: string) => {
    updateState({
      pqQuery: {
        ...state.pqQuery,
        queryName: value
      }
    });
  };

  const handleExampleClick = (exampleNumber: string) => {
    console.log(`Example ${exampleNumber} clicked`);
    showNotification(`Example ${exampleNumber} selected`, 'info');
    
    const exampleCode = examples[exampleNumber as keyof typeof examples];
    if (exampleCode) {
      handleQueryMashupChange(exampleCode);
    }
  };

  return (
    <div id="PQQuery" className="tabcontent active">
      <h3>From PQ Query</h3>
      
      <div className="pq-form-container">
        {/* Query Mashup Text Area */}
        <div className="pq-input-group">
          <label htmlFor="queryMashup" className="pq-label" title="Power Query M language code">
            query Mashup
          </label>
          <textarea
            id="queryMashup"
            className="pq-textarea"
            placeholder="Enter Power Query M language code here..."
            rows={8}
            value={state.pqQuery.queryMashup}
            onChange={(e) => handleQueryMashupChange(e.target.value)}
          />
        </div>
        
        {/* Refresh On Open Toggle */}
        <div className="pq-input-group">
          <label className="pq-toggle-container" title="Auto-refresh when opened">
            <span className="pq-toggle-label">refresh On Open</span>
            <input
              type="checkbox"
              id="refreshOnOpen"
              className="pq-toggle-input"
              checked={state.pqQuery.refreshOnOpen}
              onChange={(e) => handleRefreshOnOpenChange(e.target.checked)}
            />
            <span className="pq-toggle-slider"></span>
          </label>
        </div>
        
        {/* Query Name Input */}
        <div className="pq-input-group">
          <label htmlFor="queryName" className="pq-label" title="Query identifier">
            query Name
          </label>
          <input
            type="text"
            id="queryName"
            className="pq-input"
            value={state.pqQuery.queryName}
            placeholder="Query identifier"
            onChange={(e) => handleQueryNameChange(e.target.value)}
          />
        </div>
      </div>
      
      {/* Mashup Examples Section */}
      <div className="pq-examples-section">
        <h4 className="pq-examples-title">Mashup Example</h4>
        <div className="pq-examples-container">
          {Object.keys(examples).map((exampleNumber) => (
            <button
              key={exampleNumber}
              className="pq-example-btn"
              onClick={() => handleExampleClick(exampleNumber)}
            >
              Example {exampleNumber}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
};

export default PQQueryTab;