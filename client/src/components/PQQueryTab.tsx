import React from 'react';
import { UserSessionState } from '../types';

interface PQQueryTabProps {
  state: UserSessionState;
  updateState: (newState: Partial<UserSessionState>) => void;
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
}

const examples = {
  '1': 'let\n    Source = {1..10}   \nin \n    Source',
  '2': `let\n    Source = Web.BrowserContents("https://en.wikipedia.org/wiki/1885_Louisville_Colonels_season"),\n    ExtractedTable = Html.Table(\n        Source,\n        {\n            {"Team", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(1)"},\n            {"W", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(2)"},\n            {"L", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(3)"},\n            {"Pct.", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(4)"},\n            {"GB", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(5)"},\n            {"Home", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(6)"},\n            {"Road", "TABLE.wikitable.MLBStandingsTable > * > TR > :nth-child(7)"}\n        },\n        [RowSelector="TABLE.wikitable.MLBStandingsTable > * > TR"]\n    ),\n    PromotedHeaders = Table.PromoteHeaders(ExtractedTable, [PromoteAllScalars = true]),\n    ChangedTypes = Table.TransformColumnTypes(\n        PromotedHeaders,\n        {\n            {"Team", type text},\n            {"W", Int64.Type},\n            {"L", Int64.Type},\n            {"Pct.", type number},\n            {"GB", type text},\n            {"Home", type text},\n            {"Road", type text}\n        }\n    )\nin\n    ChangedTypes`,
  '3': 'let\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    AddedCustom = Table.AddColumn(Source, "Custom", each "Value")\nin\n    AddedCustom',
  '4': 'let\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    GroupedRows = Table.Group(Source, {"Column1"}, {{"Count", Table.RowCount, Int64.Type}})\nin\n    GroupedRows'
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