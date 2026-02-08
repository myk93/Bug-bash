import React, { createContext, useContext } from 'react';
import './App.css';
import useSessionManager from './hooks/useSessionManager';
import { useNotification } from './hooks/useNotification';
import { NotificationContainer } from './components/Notification';
import ExcelToggle from './components/ExcelToggle';
import GridTab from './components/GridTab';
import PQQueryTab from './components/PQQueryTab';
import Sidebar from './components/Sidebar';
import { TabName } from './types';

// Create context for notification
const NotificationContext = createContext<{
  showNotification: (message: string, type?: 'info' | 'success' | 'error' | 'warning') => void;
} | null>(null);

export const useNotificationContext = () => {
  const context = useContext(NotificationContext);
  if (!context) {
    throw new Error('useNotificationContext must be used within NotificationProvider');
  }
  return context;
};

const App: React.FC = () => {
  const { state, isLoading, updateState, resetSession, error } = useSessionManager();
  const { showNotification } = useNotification();

  const openTab = (tabName: TabName) => {
    // Map HTML tab names to backend state values
    const tabMapping: Record<TabName, 'grid' | 'pq-query'> = {
      'Grid': 'grid',
      'PQQuery': 'pq-query'
    };
    
    const backendTabName = tabMapping[tabName] || 'grid';
    
    // Update session state
    updateState({ activeTab: backendTabName });
    
    console.log('Tab switched to:', backendTabName);
  };

  const handleExcelToggleChange = (checked: boolean) => {
    updateState({ excelToggle: checked });

    if (checked) {
      console.log('Switched to Download mode');
    } else {
      console.log('Switched to Excel Web mode');
    }
  };

  const handleEditModeChange = (editMode: boolean) => {
    updateState({ excelWebEditMode: editMode });

    if (editMode) {
      console.log('Switched to Edit mode');
    } else {
      console.log('Switched to View mode');
    }
  };

  const handleSidebarToggle = () => {
    updateState({ sidebarCollapsed: !state.sidebarCollapsed });
  };

  const handleExportClick = () => {
    console.log('Export button clicked');
  };

  const handleResetClick = (): void => {
    resetSession();
    showNotification('Session reset successfully', 'success');
  };

  // Show loading screen while session is initializing
  if (isLoading) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        fontSize: '18px',
        color: '#333'
      }}>
        Initializing session...
      </div>
    );
  }

  // Show error if session failed to initialize
  if (error) {
    return (
      <div style={{
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        height: '100vh',
        fontSize: '18px',
        color: '#f44336'
      }}>
        Failed to initialize session: {error}
      </div>
    );
  }

  return (
    <NotificationContext.Provider value={{ showNotification }}>
      <div className="App">
        <div className="main-container">
          <div className="content-area">
            {/* Excel Toggle Switch */}
            <ExcelToggle
              excelToggle={state.excelToggle}
              onToggleChange={handleExcelToggleChange}
              onExport={handleExportClick}
              onReset={handleResetClick}
              state={state}
              showNotification={showNotification}
              onEditModeChange={handleEditModeChange}
            />

            {/* Tab Navigation */}
            <div className="tab">
              <button 
                className={`tablinks ${state.activeTab === 'grid' ? 'active' : ''}`}
                onClick={() => openTab('Grid')}
              >
                From Grid
              </button>
              <button 
                className={`tablinks ${state.activeTab === 'pq-query' ? 'active' : ''}`}
                onClick={() => openTab('PQQuery')}
              >
                From PQ Query
              </button>
            </div>

            {/* Tab Content */}
            {state.activeTab === 'grid' && (
              <GridTab
                state={state}
                updateState={updateState}
                showNotification={showNotification}
              />
            )}

            {state.activeTab === 'pq-query' && (
              <PQQueryTab
                state={state}
                updateState={updateState}
                showNotification={showNotification}
              />
            )}
          </div>

          {/* Right Sidebar */}
          <Sidebar
            isCollapsed={state.sidebarCollapsed}
            onToggle={handleSidebarToggle}
            state={state}
            updateState={updateState}
            showNotification={showNotification}
          />
        </div>

        {/* Notifications */}
        <NotificationContainer />
      </div>
    </NotificationContext.Provider>
  );
};

export default App;
