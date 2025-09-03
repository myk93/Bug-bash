
/**
 * SessionManager class handles frontend state management with session persistence
 * Integrates with backend session APIs for multi-user isolated experiences
 */
class SessionManager {
    constructor() {
        this.sessionId = null;
        this.syncTimeout = null;
        this.validationInterval = null;
        this.isInitialized = false;
        this.state = {
            activeTab: 'grid',
            excelToggle: false,
            sidebarCollapsed: false,
            fileConfigs: {
                tableName: 'Table1',
                sheetName: 'Sheet1'
            },
            docProps: {
                title: '',
                subject: '',
                keywords: '',
                createdBy: '',
                description: '',
                lastModifiedBy: '',
                category: '',
                revision: ''
            },
            pqQuery: {
                queryMashup: '',
                refreshOnOpen: false,
                queryName: 'Query1'
            },
            gridView: {
                isGridView: false,  // false = HTML Table, true = As Grid
                promoteHeaders: false,
                adjustColumnNames: false
            }
        };
        
        // Bind methods to preserve context
        this.syncToServer = this.syncToServer.bind(this);
        this.debouncedSync = this.debounce(this.syncToServer, 500);
    }

    /**
     * Initialize session - check for existing session or create new one
     */
    async init() {
        try {
            // Check sessionStorage for existing session ID
            const existingSessionId = sessionStorage.getItem('sessionId');
            
            if (existingSessionId) {
                console.log('Found existing session ID:', existingSessionId);
                const success = await this.loadExistingSession(existingSessionId);
                if (success) {
                    this.isInitialized = true;
                    this.startPeriodicSync();
                    return true;
                }
            }
            
            // Create new session if no valid existing session
            console.log('Creating new session...');
            const success = await this.createNewSession();
            if (success) {
                this.isInitialized = true;
                this.startPeriodicSync();
                return true;
            }
            
            throw new Error('Failed to initialize session');
        } catch (error) {
            console.error('Session initialization failed:', error);
            return false;
        }
    }

    /**
     * Create a new session via backend API
     */
    async createNewSession() {
        try {
            const response = await fetch('/api/session/init', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            const data = await response.json();
            
            if (data.success) {
                this.sessionId = data.sessionId;
                sessionStorage.setItem('sessionId', this.sessionId);
                
                // Load initial state from server response
                if (data.session && data.session.uiState) {
                    this.state = { ...data.session.uiState };
                }
                
                console.log('New session created:', this.sessionId);
                return true;
            } else {
                throw new Error(data.message || 'Failed to create session');
            }
        } catch (error) {
            console.error('Error creating session:', error);
            return false;
        }
    }

    /**
     * Load existing session from backend
     */
    async loadExistingSession(sessionId) {
        try {
            const response = await fetch(`/api/session/${sessionId}`);
            const data = await response.json();
            
            if (data.success) {
                this.sessionId = sessionId;
                
                // Restore state from server
                if (data.session && data.session.uiState) {
                    this.state = { ...data.session.uiState };
                }
                
                console.log('Existing session loaded:', sessionId);
                return true;
            } else {
                // Session not found or invalid, clear from storage
                sessionStorage.removeItem('sessionId');
                console.log('Session not found, will create new one');
                return false;
            }
        } catch (error) {
            console.error('Error loading session:', error);
            sessionStorage.removeItem('sessionId');
            return false;
        }
    }

    /**
     * Sync current state to server with debouncing
     */
    async syncToServer() {
        if (!this.sessionId) {
            console.warn('Cannot sync: no session ID');
            return;
        }

        try {
            const response = await fetch(`/api/session/${this.sessionId}/state`, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    uiState: this.state
                })
            });

            const data = await response.json();
            
            if (data.success) {
                console.log('State synced to server');
            } else {
                console.error('Failed to sync state:', data.message);
            }
        } catch (error) {
            console.error('Error syncing to server:', error);
        }
    }

    /**
     * Start periodic session validation (every 30 seconds)
     */
    startPeriodicSync() {
        if (this.validationInterval) {
            clearInterval(this.validationInterval);
        }
        
        this.validationInterval = setInterval(async () => {
            if (!this.sessionId) return;
            
            try {
                const response = await fetch(`/api/session/${this.sessionId}`);
                const data = await response.json();
                
                if (!data.success) {
                    console.warn('Session validation failed, reinitializing...');
                    sessionStorage.removeItem('sessionId');
                    this.sessionId = null;
                    this.isInitialized = false;
                    clearInterval(this.validationInterval);
                    await this.init();
                }
            } catch (error) {
                console.error('Session validation error:', error);
            }
        }, 30000); // 30 seconds
    }

    /**
     * Update state and trigger debounced sync
     */
    updateState(newState) {
        this.state = { ...this.state, ...newState };
        
        if (this.isInitialized) {
            this.debouncedSync();
        }
    }

    /**
     * Get current state
     */
    getState() {
        return { ...this.state };
    }

    /**
     * Utility function for debouncing
     */
    debounce(func, wait) {
        return (...args) => {
            clearTimeout(this.syncTimeout);
            this.syncTimeout = setTimeout(() => func.apply(this, args), wait);
        };
    }

    /**
     * Reset session data via backend API
     */
    async resetSession() {
        if (!this.sessionId) {
            console.warn('Cannot reset: no session ID');
            return false;
        }

        try {
            const response = await fetch(`/api/session/${this.sessionId}/reset`, {
                method: 'DELETE',
                headers: {
                    'Content-Type': 'application/json'
                }
            });

            const data = await response.json();
            
            if (data.success) {
                // Update local state with reset data
                if (data.session && data.session.uiState) {
                    this.state = { ...data.session.uiState };
                }
                
                console.log('Session reset successfully');
                return true;
            } else {
                console.error('Failed to reset session:', data.message);
                return false;
            }
        } catch (error) {
            console.error('Error resetting session:', error);
            return false;
        }
    }

    /**
     * Cleanup when page unloads
     */
    cleanup() {
        if (this.syncTimeout) {
            clearTimeout(this.syncTimeout);
        }
        if (this.validationInterval) {
            clearInterval(this.validationInterval);
        }
    }
}

// Global session manager instance
const sessionManager = new SessionManager();

/**
 * Enhanced openTab function with session state management
 */
function openTab(evt, tabName) {
    var i, tabcontent, tablinks;
    
    // Hide all tab content
    tabcontent = document.getElementsByClassName("tabcontent");
    for (i = 0; i < tabcontent.length; i++) {
        tabcontent[i].style.display = "none";
    }
    
    // Remove active class from all tab links
    tablinks = document.getElementsByClassName("tablinks");
    for (i = 0; i < tablinks.length; i++) {
        tablinks[i].className = tablinks[i].className.replace(" active", "");
    }
    
    // Show selected tab and mark as active
    document.getElementById(tabName).style.display = "block";
    evt.currentTarget.className += " active";
    
    // Map HTML tab names to backend state values
    const tabMapping = {
        'Grid': 'grid',
        'HTMLTable': 'table', 
        'PQQuery': 'pq-query'
    };
    
    const backendTabName = tabMapping[tabName] || 'grid';
    
    // Update session state
    sessionManager.updateState({ activeTab: backendTabName });
    
    console.log('Tab switched to:', backendTabName);
}

/**
 * Function to handle "Open in Excel for the web" action
 */
function handleOpenInExcel() {
    console.log('Opening in Excel for the web...');
    // Add your logic here for opening in Excel for the web
    // For example: window.open('https://office.live.com/start/Excel.aspx', '_blank');
    showNotification('Ready to open in Excel for the web');
}

/**
 * Function to handle "Download workbook" action
 */
function handleDownloadWorkbook() {
    console.log('Preparing workbook download...');
    // Add your logic here for downloading the workbook
    // For example: triggerDownload();
    showNotification('Ready to download workbook');
}

/**
 * Generate filename for Excel export
 * Creates a filename in the format: {tabName}_{dateTime}.xlsx
 * @returns {string} Generated filename
 */
function generateFileName() {
    try {
        // Get the current active tab name
        const activeTab = sessionManager.getState().activeTab;
        
        // Map backend tab names to user-friendly names
        const tabNameMapping = {
            'grid': 'Grid',
            'pq-query': 'PQQuery',
            'table': 'HTMLTable'
        };
        
        const tabName = tabNameMapping[activeTab] || 'Export';
        
        // Get current date and time in readable format
        const now = new Date();
        const year = now.getFullYear();
        const month = String(now.getMonth() + 1).padStart(2, '0');
        const day = String(now.getDate()).padStart(2, '0');
        const hours = String(now.getHours()).padStart(2, '0');
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const seconds = String(now.getSeconds()).padStart(2, '0');
        
        const dateTime = `${year}-${month}-${day}_${hours}-${minutes}-${seconds}`;
        
        // Return filename in format: {tabName}_{dateTime}.xlsx
        const filename = `${tabName}_${dateTime}.xlsx`;
        
        console.log('Generated filename:', filename);
        return filename;
        
    } catch (error) {
        console.error('Error generating filename:', error);
        // Fallback filename if generation fails
        const fallbackDateTime = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
        return `Export_${fallbackDateTime}.xlsx`;
    }
}

/**
 * Handle Excel toggle functionality
 * Checks toggle state and calls appropriate function (Excel Web or Download)
 * @param {Blob} blob - The workbook blob to process
 */
async function handleExcelToggle(blob) {
    try {
        if (!blob) {
            throw new Error('No workbook blob provided');
        }
        
        // Get the Excel toggle element
        const excelToggle = document.getElementById('excelToggle');
        if (!excelToggle) {
            throw new Error('Excel toggle element not found');
        }
        
        // Generate filename
        const filename = generateFileName();
        
        // Check toggle state - checked means Download mode, unchecked means Excel Web mode
        const isDownloadMode = excelToggle.checked;
        
        console.log('Excel toggle state:', isDownloadMode ? 'Download' : 'Excel Web');
        console.log('Processing with filename:', filename);
        
        if (isDownloadMode) {
            // Download mode - trigger browser download
            console.log('Triggering browser download...');
            await downloadWorkbookBlob(blob, filename);
            showNotification(`Workbook downloaded: ${filename}`, 'success');
        } else {
            // Excel Web mode - open in Excel Web
            console.log('Opening in Excel Web...');
            await openInExcelWeb(blob, filename);
            showNotification(`Opened in Excel Web: ${filename}`, 'success');
        }
        
    } catch (error) {
        console.error('Error handling Excel toggle:', error);
        showNotification(`Excel operation failed: ${error.message}`, 'error');
        throw error;
    }
}

/**
 * Function to handle "Export" action
 */
async function handleExportData() {
    try {
        console.log('Exporting data...');
        showNotification('Preparing export...', 'info');
        
        // Get the current active tab from session manager
        const activeTab = sessionManager.getState().activeTab;
        console.log('Active tab for export:', activeTab);
        
        let blob = null;
        
        console.log('Generating workbook for active tab:', activeTab);
        if (activeTab === 'pq-query') {
            // For PQQuery tab: call generateQueryWorkbook to get the blob
            console.log('Generating workbook for PQQuery tab...');
            blob = await generateQueryWorkbook();
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
            await handleExcelToggle(blob);
        } else {
            throw new Error('Failed to generate workbook blob');
        }
        
    } catch (error) {
        console.error('Error exporting data:', error);
        showNotification(`Export failed: ${error.message}`, 'error');
    }
}

/**
 * Function to show notification
 */
function showNotification(message) {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('excel-notification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'excel-notification';
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background-color: #217346;
            color: white;
            padding: 12px 20px;
            border-radius: 6px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            z-index: 1001;
            font-size: 14px;
            font-weight: 500;
            opacity: 0;
            transition: opacity 0.3s ease;
        `;
        document.body.appendChild(notification);
    }
    
    notification.textContent = message;
    notification.style.opacity = '1';
    
    // Hide notification after 3 seconds
    setTimeout(() => {
        notification.style.opacity = '0';
    }, 3000);
}

/**
 * Restore UI state from session manager
 */
function restoreUIState() {
    const state = sessionManager.getState();
    
    // Restore active tab
    const tabMapping = {
        'grid': 'Grid',
        'table': 'HTMLTable',
        'pq-query': 'PQQuery'
    };
    
    const htmlTabName = tabMapping[state.activeTab] || 'Grid';
    
    // Find and click the appropriate tab using data attributes
    const tablinks = document.getElementsByClassName("tablinks");
    for (let i = 0; i < tablinks.length; i++) {
        const button = tablinks[i];
        if (button.dataset.tab === htmlTabName) {
            // Create a mock event object for the openTab function
            const mockEvent = { currentTarget: button };
            openTab(mockEvent, htmlTabName);
            break;
        }
    }
    
    // Restore Excel toggle state
    const excelToggle = document.getElementById('excelToggle');
    if (excelToggle) {
        excelToggle.checked = state.excelToggle;
        
        // Trigger the change event to ensure proper state with safety checks
        if (state.excelToggle) {
            if (typeof handleDownloadWorkbook === 'function') {
                handleDownloadWorkbook();
            } else {
                console.warn('handleDownloadWorkbook function not available');
            }
        } else {
            if (typeof handleOpenInExcel === 'function') {
                handleOpenInExcel();
            } else {
                console.warn('handleOpenInExcel function not available');
            }
        }
    }
    
    // Restore sidebar state
    if (state.sidebarCollapsed) {
        toggleSidebar(true);
    }
    
    // Restore file configs
    if (state.fileConfigs) {
        const tableNameInput = document.getElementById('tableName');
        const sheetNameInput = document.getElementById('sheetName');
        
        if (tableNameInput) tableNameInput.value = state.fileConfigs.tableName || 'Table1';
        if (sheetNameInput) sheetNameInput.value = state.fileConfigs.sheetName || 'Sheet1';
    }
    
    // Restore document properties
    if (state.docProps) {
        Object.keys(state.docProps).forEach(key => {
            const element = document.getElementById(`doc${key.charAt(0).toUpperCase() + key.slice(1)}`);
            if (element && state.docProps[key]) {
                element.value = state.docProps[key];
            }
        });
    }
    
    // Restore PQ Query state
    if (state.pqQuery) {
        const queryMashupElement = document.getElementById('queryMashup');
        const refreshOnOpenElement = document.getElementById('refreshOnOpen');
        const queryNameElement = document.getElementById('queryName');
        
        if (queryMashupElement) queryMashupElement.value = state.pqQuery.queryMashup || '';
        if (refreshOnOpenElement) refreshOnOpenElement.checked = state.pqQuery.refreshOnOpen || false;
        if (queryNameElement) queryNameElement.value = state.pqQuery.queryName || 'Query1';
    }
    
    // Restore Grid View state
    if (state.gridView) {
        const gridViewToggle = document.getElementById('gridViewToggle');
        const gridOptions = document.getElementById('gridOptions');
        const promoteHeadersElement = document.getElementById('promoteHeaders');
        const adjustColumnNamesElement = document.getElementById('adjustColumnNames');
        
        if (gridViewToggle) {
            gridViewToggle.checked = state.gridView.isGridView || false;
            
            // Show/hide grid options based on toggle state
            if (gridOptions) {
                gridOptions.style.display = state.gridView.isGridView ? 'flex' : 'none';
            }
        }
        
        if (promoteHeadersElement) promoteHeadersElement.checked = state.gridView.promoteHeaders || false;
        if (adjustColumnNamesElement) adjustColumnNamesElement.checked = state.gridView.adjustColumnNames || false;
    }
    
    console.log('UI state restored:', state);
}

/**
 * Toggle sidebar collapsed state
 */
function toggleSidebar(forceState = null) {
    const sidebar = document.getElementById('rightSidebar');
    const isCurrentlyCollapsed = sidebar.classList.contains('collapsed');
    const shouldCollapse = forceState !== null ? forceState : !isCurrentlyCollapsed;
    
    if (shouldCollapse) {
        sidebar.classList.add('collapsed');
    } else {
        sidebar.classList.remove('collapsed');
    }
    
    // Update session state
    sessionManager.updateState({ sidebarCollapsed: shouldCollapse });
}

/**
 * Update file configs in session
 */
function updateFileConfigs() {
    const fileConfigs = getFileConfigs();
    sessionManager.updateState({ fileConfigs });
}

/**
 * Update document properties in session
 */
function updateDocProps() {
    const docProps = getDocumentProperties();
    sessionManager.updateState({ docProps });
}

/**
 * Update PQ Query data in session
 */
function updatePQQuery() {
    const pqQuery = getPQQueryData();
    sessionManager.updateState({ pqQuery });
}

/**
 * Get current PQ Query data
 */
function getPQQueryData() {
    const queryMashup = document.getElementById('queryMashup')?.value || '';
    const refreshOnOpen = document.getElementById('refreshOnOpen')?.checked || false;
    const queryName = document.getElementById('queryName')?.value || 'Query1';
    
    return {
        queryMashup,
        refreshOnOpen,
        queryName
    };
}

/**
 * Get active tab data based on session manager state
 * Returns structured data for the currently active tab
 *
 * This function determines which tab is currently active using sessionManager.getState().activeTab
 * and collects the appropriate data based on the active tab:
 *
 * - For 'pq-query' tab: Collects QueryInfo data using getPQQueryData() and maps it to the
 *   QueryInfo interface structure with queryMashup, refreshOnOpen, and queryName properties
 * - For 'grid' tab: Currently returns null for gridData (logic to be implemented later)
 *
 * @returns {Object} Object containing:
 *   - activeTab: string - The currently active tab ('pq-query', 'grid', etc.)
 *   - queryInfo: QueryInfo|null - Query data if PQQuery tab is active, null otherwise
 *   - gridData: any|null - Grid data if Grid tab is active (currently always null)
 *
 * QueryInfo interface structure:
 * {
 *   queryMashup: string,        // Power Query M language code
 *   refreshOnOpen: boolean,     // Auto-refresh when opened
 *   queryName?: string          // Query identifier (default: "Query1")
 * }
 */
function getActiveTabData() {
    // Get current active tab from session manager
    const activeTab = sessionManager.getState().activeTab;
    
    let queryInfo = null;
    let gridData = null;
    console.log('Active tab data:', activeTab);
    // Collect data based on active tab
    if (activeTab === 'pq-query') {
        // For PQQuery tab, collect QueryInfo data using existing function
        const pqData = getPQQueryData();
        
        // Map to QueryInfo interface structure
        queryInfo = {
            queryMashup: pqData.queryMashup,
            refreshOnOpen: pqData.refreshOnOpen,
            queryName: pqData.queryName
        };
    } else if (activeTab === 'grid') {
        // For Grid tab, leave logic empty for now as requested
        gridData = null;
    }
    
    // Return structured object with active tab and appropriate data
    return {
        activeTab: activeTab,
        queryInfo: queryInfo,
        gridData: gridData
    };
}

/**
 * Get current Grid View data
 */
function getGridViewData() {
    const gridViewToggle = document.getElementById('gridViewToggle')?.checked || false;
    const promoteHeaders = document.getElementById('promoteHeaders')?.checked || false;
    const adjustColumnNames = document.getElementById('adjustColumnNames')?.checked || false;
    
    return {
        isGridView: gridViewToggle,
        promoteHeaders,
        adjustColumnNames
    };
}

/**
 * Update Grid View data in session
 */
function updateGridView() {
    const gridView = getGridViewData();
    sessionManager.updateState({ gridView });
}

/**
 * Set PQ Query data
 */
function setPQQueryData(pqData) {
    const queryMashupElement = document.getElementById('queryMashup');
    const refreshOnOpenElement = document.getElementById('refreshOnOpen');
    const queryNameElement = document.getElementById('queryName');
    
    if (queryMashupElement && pqData.queryMashup !== undefined) {
        queryMashupElement.value = pqData.queryMashup;
    }
    if (refreshOnOpenElement && pqData.refreshOnOpen !== undefined) {
        refreshOnOpenElement.checked = pqData.refreshOnOpen;
    }
    if (queryNameElement && pqData.queryName !== undefined) {
        queryNameElement.value = pqData.queryName;
    }
}

/**
 * Enhanced Excel toggle functionality with session persistence
 */
document.addEventListener('DOMContentLoaded', async function() {
    // Initialize session manager
    const initSuccess = await sessionManager.init();
    if (!initSuccess) {
        console.error('Failed to initialize session management');
        return;
    }
    
    // Add event listeners to tab buttons
    const tablinks = document.getElementsByClassName("tablinks");
    for (let i = 0; i < tablinks.length; i++) {
        const button = tablinks[i];
        button.addEventListener('click', function(event) {
            const tabName = this.dataset.tab;
            if (tabName) {
                openTab(event, tabName);
            }
        });
    }
    
    // Restore UI state after session is initialized and event listeners are added
    restoreUIState();
    
    const excelToggle = document.getElementById('excelToggle');
    
    // Handle toggle change with session persistence
    excelToggle.addEventListener('change', function() {
        const isChecked = this.checked;
        
        // Update session state
        sessionManager.updateState({ excelToggle: isChecked });
        
        if (isChecked) {
            // Download workbook mode
            handleDownloadWorkbook();
        } else {
            // Open in Excel for the web mode
            handleOpenInExcel();
        }
        
        console.log('Excel toggle state:', isChecked);
    });
    
    // Functions are now defined globally above - no need to redefine them here
    
    // Reset functionality
    async function resetUserSession() {
        // Show confirmation dialog
        const confirmed = confirm(
            'Reset your workspace? This will clear all your data.\n\n' +
            'This action cannot be undone. Your session will be reset to its default state.'
        );
        
        if (!confirmed) {
            return;
        }
        
        try {
            // Show loading state
            const resetButton = document.getElementById('resetButton');
            const originalText = resetButton.innerHTML;
            resetButton.innerHTML = '<span class="reset-icon">⏳</span>Resetting...';
            resetButton.disabled = true;
            
            // Call the session manager reset method
            const success = await sessionManager.resetSession();
            
            if (success) {
                // Show success message
                showNotification('Workspace reset successfully');
                
                // Reload the page to show clean state
                setTimeout(() => {
                    window.location.reload();
                }, 1000);
            } else {
                // Show error message
                showNotification('Failed to reset workspace. Please try again.');
                
                // Restore button state
                resetButton.innerHTML = originalText;
                resetButton.disabled = false;
            }
        } catch (error) {
            console.error('Error during reset:', error);
            showNotification('An error occurred while resetting. Please try again.');
            
            // Restore button state
            const resetButton = document.getElementById('resetButton');
            resetButton.innerHTML = '<span class="reset-icon">⚠️</span>Reset Workspace';
            resetButton.disabled = false;
        }
    }
    
    // Add event listener for reset button
    const resetButton = document.getElementById('resetButton');
    if (resetButton) {
        resetButton.addEventListener('click', resetUserSession);
    }
    
    // Add event listener for export button
    const exportButton = document.getElementById('exportButton');
    if (exportButton) {
        exportButton.addEventListener('click', handleExportData);
    }
    
    // Make reset function available globally
    window.resetUserSession = resetUserSession;
    
    // Cleanup on page unload
    window.addEventListener('beforeunload', () => {
        sessionManager.cleanup();
    });
    
    // File upload functionality
    initializeFileUpload();
    
    // Initialize sidebar functionality
    initializeSidebar();
    
    // Initialize form persistence
    initializeFormPersistence();
    
    // Initialize PQ Query functionality
    initializePQQuery();
    
    // Initialize Grid functionality
    initializeEditableGrid();
    
    // Initialize Grid View functionality
    initializeGridView();
});

/**
 * Initialize sidebar toggle functionality
 */
function initializeSidebar() {
    const sidebarToggle = document.getElementById('sidebarToggle');
    
    if (!sidebarToggle) {
        console.warn('Sidebar toggle element not found');
        return;
    }
    
    // Add click event listener for sidebar toggle
    sidebarToggle.addEventListener('click', () => {
        toggleSidebar();
    });
}

/**
 * Initialize form persistence for file configs and doc props
 */
function initializeFormPersistence() {
    // Add event listeners for file config inputs
    const tableNameInput = document.getElementById('tableName');
    const sheetNameInput = document.getElementById('sheetName');
    
    if (tableNameInput) {
        tableNameInput.addEventListener('input', updateFileConfigs);
        tableNameInput.addEventListener('change', updateFileConfigs);
    }
    
    if (sheetNameInput) {
        sheetNameInput.addEventListener('input', updateFileConfigs);
        sheetNameInput.addEventListener('change', updateFileConfigs);
    }
    
    // Add event listeners for document properties inputs
    const docPropIds = [
        'docTitle', 'docSubject', 'docKeywords', 'docCreatedBy',
        'docDescription', 'docLastModifiedBy', 'docCategory', 'docRevision'
    ];
    
    docPropIds.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('input', updateDocProps);
            element.addEventListener('change', updateDocProps);
        }
    });
}

/**
 * Initialize PQ Query functionality
 */
function initializePQQuery() {
    // Add event listeners for PQ Query inputs
    const queryMashupElement = document.getElementById('queryMashup');
    const refreshOnOpenElement = document.getElementById('refreshOnOpen');
    const queryNameElement = document.getElementById('queryName');
    
    if (queryMashupElement) {
        queryMashupElement.addEventListener('input', updatePQQuery);
        queryMashupElement.addEventListener('change', updatePQQuery);
    }
    
    if (refreshOnOpenElement) {
        refreshOnOpenElement.addEventListener('change', updatePQQuery);
    }
    
    if (queryNameElement) {
        queryNameElement.addEventListener('input', updatePQQuery);
        queryNameElement.addEventListener('change', updatePQQuery);
    }
    
    // Add event listeners for example buttons
    const exampleButtons = document.querySelectorAll('.pq-example-btn');
    exampleButtons.forEach(button => {
        button.addEventListener('click', function() {
            const exampleNumber = this.dataset.example;
            handleExampleClick(exampleNumber);
        });
    });
}

/**
 * Handle example button clicks
 */
function handleExampleClick(exampleNumber) {
    console.log(`Example ${exampleNumber} clicked`);
    showNotification(`Example ${exampleNumber} selected`, 'info');
    
    // You can add specific example code here for each button
    const examples = {
        '1': '// Example 1: Basic table query\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content]\nin\n    Source',
        '2': '// Example 2: Filter and transform\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    FilteredRows = Table.SelectRows(Source, each ([Column1] <> null))\nin\n    FilteredRows',
        '3': '// Example 3: Add custom column\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    AddedCustom = Table.AddColumn(Source, "Custom", each "Value")\nin\n    AddedCustom',
        '4': '// Example 4: Group and aggregate\nlet\n    Source = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],\n    GroupedRows = Table.Group(Source, {"Column1"}, {{"Count", Table.RowCount, Int64.Type}})\nin\n    GroupedRows'
    };
    
    const exampleCode = examples[exampleNumber];
    if (exampleCode) {
        const queryMashupElement = document.getElementById('queryMashup');
        if (queryMashupElement) {
            queryMashupElement.value = exampleCode;
            // Trigger the change event to save to session
            updatePQQuery();
        }
    }
}

/**
 * Initialize drag and drop file upload functionality
 */
function initializeFileUpload() {
    const dragDropArea = document.getElementById('dragDropArea');
    const fileInput = document.getElementById('fileInput');
    
    if (!dragDropArea || !fileInput) {
        console.warn('File upload elements not found');
        return;
    }
    
    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dragDropArea.addEventListener(eventName, preventDefaults, false);
        document.body.addEventListener(eventName, preventDefaults, false);
    });
    
    // Highlight drop area when item is dragged over it
    ['dragenter', 'dragover'].forEach(eventName => {
        dragDropArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        dragDropArea.addEventListener(eventName, unhighlight, false);
    });
    
    // Handle dropped files
    dragDropArea.addEventListener('drop', handleDrop, false);
    
    // Handle click to browse files
    dragDropArea.addEventListener('click', () => {
        fileInput.click();
    });
    
    // Handle file input change
    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    const dragDropArea = document.getElementById('dragDropArea');
    dragDropArea.classList.add('dragover');
}

function unhighlight(e) {
    const dragDropArea = document.getElementById('dragDropArea');
    dragDropArea.classList.remove('dragover');
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
}

function handleFiles(files) {
    [...files].forEach(handleFile);
}

function handleFile(file) {
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
    
    // Update the drag and drop area to show uploaded file
    updateUploadDisplay(file);
    
    console.log('File uploaded:', file.name, 'Size:', file.size, 'Type:', file.type);
}

function updateUploadDisplay(file) {
    const dragDropArea = document.getElementById('dragDropArea');
    const dragDropContent = dragDropArea.querySelector('.drag-drop-content');
    
    dragDropContent.innerHTML = `
        <span class="upload-icon">✅</span>
        <p><strong>${file.name}</strong></p>
        <p class="upload-hint">File uploaded successfully</p>
        <p class="upload-hint">Click to upload another file</p>
    `;
    
    dragDropArea.style.borderColor = '#4CAF50';
    dragDropArea.style.backgroundColor = '#f0f8f0';
}

/**
 * Enhanced notification function with different types
 */
function showNotification(message, type = 'info') {
    // Create notification element if it doesn't exist
    let notification = document.getElementById('excel-notification');
    if (!notification) {
        notification = document.createElement('div');
        notification.id = 'excel-notification';
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 12px 20px;
            border-radius: 6px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
            z-index: 1001;
            font-size: 14px;
            font-weight: 500;
            opacity: 0;
            transition: opacity 0.3s ease;
        `;
        document.body.appendChild(notification);
    }
    
    // Set colors based on type
    let backgroundColor, color;
    switch(type) {
        case 'success':
            backgroundColor = '#4CAF50';
            color = 'white';
            break;
        case 'error':
            backgroundColor = '#f44336';
            color = 'white';
            break;
        case 'warning':
            backgroundColor = '#ff9800';
            color = 'white';
            break;
        default:
            backgroundColor = '#217346';
            color = 'white';
    }
    
    notification.style.backgroundColor = backgroundColor;
    notification.style.color = color;
    notification.textContent = message;
    notification.style.opacity = '1';
    
    // Hide notification after 4 seconds
    setTimeout(() => {
        notification.style.opacity = '0';
    }, 4000);
}

/**
 * Get current file configs
 */
function getFileConfigs() {
    const tableName = document.getElementById('tableName')?.value || 'Table1';
    const sheetName = document.getElementById('sheetName')?.value || 'Sheet1';
    
    return {
        tableName,
        sheetName
    };
}

/**
 * Get current document properties
 */
function getDocumentProperties() {
    return {
        title: document.getElementById('docTitle')?.value || '',
        subject: document.getElementById('docSubject')?.value || '',
        keywords: document.getElementById('docKeywords')?.value || '',
        createdBy: document.getElementById('docCreatedBy')?.value || '',
        description: document.getElementById('docDescription')?.value || '',
        lastModifiedBy: document.getElementById('docLastModifiedBy')?.value || '',
        category: document.getElementById('docCategory')?.value || '',
        revision: document.getElementById('docRevision')?.value || ''
    };
}

/**
 * Collect FileConfig data from the config tab fields
 * Maps data to the FileConfig interface structure required by @microsoft/connected-workbooks
 * Only includes properties that have actual user-entered data (not empty strings)
 */
function collectFileConfigData() {
    const fileConfig = {};
    
    // Get document properties using existing helper function
    const docProps = getDocumentProperties();
    
    // Filter out empty document properties
    const filteredDocProps = {};
    Object.keys(docProps).forEach(key => {
        if (docProps[key] && docProps[key].trim() !== '') {
            filteredDocProps[key] = docProps[key].trim();
        }
    });
    
    // Only add docProps if it has any non-empty values
    if (Object.keys(filteredDocProps).length > 0) {
        fileConfig.docProps = filteredDocProps;
    }
    
    // Get file configs using existing helper function
    const fileConfigs = getFileConfigs();
    
    // Create TempleteSettings object and filter out empty values
    const templeteSettings = {};
    if (fileConfigs.tableName && fileConfigs.tableName.trim() !== '') {
        templeteSettings.tableName = fileConfigs.tableName.trim();
    }
    if (fileConfigs.sheetName && fileConfigs.sheetName.trim() !== '') {
        templeteSettings.sheetName = fileConfigs.sheetName.trim();
    }
    
    // Only add TempleteSettings if it has any non-empty values
    if (Object.keys(templeteSettings).length > 0) {
        fileConfig.TempleteSettings = templeteSettings;
    }
    
    // templateFile is left undefined for now as specified
    // fileConfig.templateFile = undefined;
    
    return fileConfig;
}

/**
 * Set document properties
 */
function setDocumentProperties(props) {
    Object.keys(props).forEach(key => {
        const element = document.getElementById(`doc${key.charAt(0).toUpperCase() + key.slice(1)}`);
        if (element && props[key]) {
            element.value = props[key];
        }
    });
}

/**
 * Initialize the editable grid functionality
 */
function initializeEditableGrid() {
    // Grid state
    let gridData = [];
    let currentRows = 3;
    let currentCols = 3;
    
    // Initialize with default 3x3 grid
    createGrid(currentRows, currentCols);
    
    // Add event listeners
    const updateBtn = document.getElementById('updateGrid');
    const clearBtn = document.getElementById('clearGrid');
    const importBtn = document.getElementById('importGrid');
    const rowCountInput = document.getElementById('rowCount');
    const colCountInput = document.getElementById('colCount');
    
    if (updateBtn) {
        updateBtn.addEventListener('click', handleUpdateGrid);
    }
    
    if (clearBtn) {
        clearBtn.addEventListener('click', handleClearGrid);
    }
    
    if (importBtn) {
        importBtn.addEventListener('click', handleImportGrid);
    }
    
    if (rowCountInput) {
        rowCountInput.addEventListener('input', validateGridSize);
    }
    
    if (colCountInput) {
        colCountInput.addEventListener('input', validateGridSize);
    }
    
    /**
     * Create the editable grid with specified dimensions
     */
    function createGrid(rows, cols) {
        const gridBody = document.getElementById('gridBody');
        const gridInfo = document.getElementById('gridInfo');
        
        if (!gridBody) return;
        
        // Clear existing grid
        gridBody.innerHTML = '';
        
        // Preserve existing data or initialize new data
        if (gridData.length === 0 || gridData.length !== rows) {
            gridData = Array(rows).fill().map(() => Array(cols).fill(''));
        } else {
            // Adjust existing data to new dimensions
            gridData = gridData.slice(0, rows);
            for (let i = 0; i < rows; i++) {
                if (!gridData[i]) {
                    gridData[i] = Array(cols).fill('');
                } else {
                    gridData[i] = gridData[i].slice(0, cols);
                    while (gridData[i].length < cols) {
                        gridData[i].push('');
                    }
                }
            }
        }
        
        // Create grid rows and cells
        for (let i = 0; i < rows; i++) {
            const row = document.createElement('tr');
            
            for (let j = 0; j < cols; j++) {
                const cell = document.createElement('td');
                const textarea = document.createElement('textarea');
                
                textarea.className = 'grid-cell';
                textarea.value = gridData[i][j] || '';
                textarea.placeholder = `R${i + 1}C${j + 1}`;
                textarea.rows = 1;
                
                // Add event listeners for cell interactions
                textarea.addEventListener('input', function() {
                    gridData[i][j] = this.value;
                    autoResizeTextarea(this);
                    updateGridInSession();
                });
                
                textarea.addEventListener('keydown', function(e) {
                    handleCellNavigation(e, i, j, rows, cols);
                });
                
                textarea.addEventListener('focus', function() {
                    this.select();
                });
                
                // Auto-resize on load
                autoResizeTextarea(textarea);
                
                cell.appendChild(textarea);
                row.appendChild(cell);
            }
            
            gridBody.appendChild(row);
        }
        
        // Update grid info
        if (gridInfo) {
            gridInfo.textContent = `${rows} rows × ${cols} columns`;
        }
        
        // Update current dimensions
        currentRows = rows;
        currentCols = cols;
        
        console.log('Grid created:', rows, 'x', cols);
    }
    
    /**
     * Handle grid update button click
     */
    function handleUpdateGrid() {
        const newRows = parseInt(rowCountInput.value) || 3;
        const newCols = parseInt(colCountInput.value) || 3;
        
        if (newRows < 1 || newRows > 20 || newCols < 1 || newCols > 20) {
            showNotification('Grid size must be between 1x1 and 20x20', 'error');
            return;
        }
        
        createGrid(newRows, newCols);
        showNotification(`Grid updated to ${newRows}×${newCols}`, 'success');
    }
    
    /**
     * Handle clear grid button click
     */
    function handleClearGrid() {
        const confirmed = confirm('Clear all grid data? This action cannot be undone.');
        if (!confirmed) return;
        
        gridData = Array(currentRows).fill().map(() => Array(currentCols).fill(''));
        
        // Clear all textareas
        const cells = document.querySelectorAll('.grid-cell');
        cells.forEach(cell => {
            cell.value = '';
            autoResizeTextarea(cell);
        });
        
        showNotification('Grid cleared successfully', 'success');
        updateGridInSession();
    }
    
    /**
     * Handle import grid data
     */
    function handleImportGrid() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.txt';
        
        input.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const content = e.target.result;
                    const rows = content.split('\n').filter(row => row.trim());
                    const importedData = rows.map(row => {
                        // Simple CSV parsing (handles quoted values)
                        const cells = [];
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
                        
                        // Update input values
                        rowCountInput.value = newRows;
                        colCountInput.value = newCols;
                        
                        // Create new grid
                        createGrid(newRows, newCols);
                        
                        // Fill with imported data
                        for (let i = 0; i < newRows; i++) {
                            for (let j = 0; j < newCols; j++) {
                                if (importedData[i] && importedData[i][j] !== undefined) {
                                    gridData[i][j] = importedData[i][j];
                                    const cell = document.querySelector(`tr:nth-child(${i + 1}) td:nth-child(${j + 1}) .grid-cell`);
                                    if (cell) {
                                        cell.value = importedData[i][j];
                                        autoResizeTextarea(cell);
                                    }
                                }
                            }
                        }
                        
                        showNotification('Grid data imported successfully', 'success');
                        updateGridInSession();
                    }
                } catch (error) {
                    console.error('Import error:', error);
                    showNotification('Error importing file. Please check file format.', 'error');
                }
            };
            reader.readAsText(file);
        });
        
        input.click();
    }
    
    /**
     * Validate grid size inputs
     */
    function validateGridSize() {
        const rows = parseInt(rowCountInput.value);
        const cols = parseInt(colCountInput.value);
        
        if (rows < 1 || rows > 20) {
            rowCountInput.style.borderColor = '#f44336';
        } else {
            rowCountInput.style.borderColor = '#ddd';
        }
        
        if (cols < 1 || cols > 20) {
            colCountInput.style.borderColor = '#f44336';
        } else {
            colCountInput.style.borderColor = '#ddd';
        }
    }
    
    /**
     * Handle keyboard navigation between cells
     */
    function handleCellNavigation(e, row, col, totalRows, totalCols) {
        let targetRow = row;
        let targetCol = col;
        
        switch (e.key) {
            case 'Tab':
                e.preventDefault();
                if (e.shiftKey) {
                    // Move to previous cell
                    targetCol--;
                    if (targetCol < 0) {
                        targetCol = totalCols - 1;
                        targetRow--;
                        if (targetRow < 0) {
                            targetRow = totalRows - 1;
                        }
                    }
                } else {
                    // Move to next cell
                    targetCol++;
                    if (targetCol >= totalCols) {
                        targetCol = 0;
                        targetRow++;
                        if (targetRow >= totalRows) {
                            targetRow = 0;
                        }
                    }
                }
                break;
                
            case 'ArrowUp':
                if (e.ctrlKey) {
                    e.preventDefault();
                    targetRow = Math.max(0, targetRow - 1);
                }
                break;
                
            case 'ArrowDown':
                if (e.ctrlKey) {
                    e.preventDefault();
                    targetRow = Math.min(totalRows - 1, targetRow + 1);
                }
                break;
                
            case 'ArrowLeft':
                if (e.ctrlKey) {
                    e.preventDefault();
                    targetCol = Math.max(0, targetCol - 1);
                }
                break;
                
            case 'ArrowRight':
                if (e.ctrlKey) {
                    e.preventDefault();
                    targetCol = Math.min(totalCols - 1, targetCol + 1);
                }
                break;
                
            case 'Enter':
                e.preventDefault();
                targetRow = Math.min(totalRows - 1, targetRow + 1);
                break;
        }
        
        // Focus target cell if different from current
        if (targetRow !== row || targetCol !== col) {
            const targetCell = document.querySelector(`tr:nth-child(${targetRow + 1}) td:nth-child(${targetCol + 1}) .grid-cell`);
            if (targetCell) {
                targetCell.focus();
            }
        }
    }
    
    /**
     * Auto-resize textarea based on content
     */
    function autoResizeTextarea(textarea) {
        textarea.style.height = 'auto';
        const scrollHeight = textarea.scrollHeight;
        const minHeight = 40;
        textarea.style.height = Math.max(minHeight, scrollHeight) + 'px';
    }
    
    /**
     * Update grid data in session
     */
    function updateGridInSession() {
        if (sessionManager && sessionManager.isInitialized) {
            sessionManager.updateState({
                gridData: {
                    data: gridData,
                    rows: currentRows,
                    cols: currentCols
                }
            });
        }
    }
    
    /**
     * Load grid data from session
     */
    function loadGridFromSession() {
        const state = sessionManager.getState();
        if (state.gridData && state.gridData.data) {
            const { data, rows, cols } = state.gridData;
            if (rows && cols) {
                currentRows = rows;
                currentCols = cols;
                gridData = data;
                
                // Update input values
                rowCountInput.value = rows;
                colCountInput.value = cols;
                
                // Recreate grid
                createGrid(rows, cols);
                
                console.log('Grid loaded from session:', rows, 'x', cols);
            }
        }
    }
    
    // Load grid from session if available
    if (sessionManager && sessionManager.isInitialized) {
        loadGridFromSession();
    }
}

/**
 * Initialize Grid View functionality
 */
function initializeGridView() {
    const gridViewToggle = document.getElementById('gridViewToggle');
    const gridOptions = document.getElementById('gridOptions');
    const promoteHeadersElement = document.getElementById('promoteHeaders');
    const adjustColumnNamesElement = document.getElementById('adjustColumnNames');
    
    if (!gridViewToggle || !gridOptions) {
        console.warn('Grid view elements not found');
        return;
    }
    
    // Handle grid view toggle change
    gridViewToggle.addEventListener('change', function() {
        const isGridView = this.checked;
        
        // Show/hide grid options based on toggle state
        if (isGridView) {
            gridOptions.style.display = 'flex';
            showNotification('Switched to Grid view', 'info');
        } else {
            gridOptions.style.display = 'none';
            showNotification('Switched to HTML Table view', 'info');
        }
        
        // Update session state
        updateGridView();
        
        console.log('Grid view toggled:', isGridView ? 'As Grid' : 'HTML Table');
    });
    
    // Handle promote headers checkbox change
    if (promoteHeadersElement) {
        promoteHeadersElement.addEventListener('change', function() {
            updateGridView();
            console.log('Promote headers:', this.checked);
            
            if (this.checked) {
                showNotification('Headers will be promoted from first row', 'info');
            }
        });
    }
    
    // Handle adjust column names checkbox change
    if (adjustColumnNamesElement) {
        adjustColumnNamesElement.addEventListener('change', function() {
            updateGridView();
            console.log('Adjust column names:', this.checked);
            
            if (this.checked) {
                showNotification('Column names will be adjusted for duplicates/invalid names', 'info');
            }
        });
    }
    
    console.log('Grid view functionality initialized');
}

/**
 * Generate query workbook using the server-side API
 *
 * This function collects data from the currently active tab and file configurations,
 * then calls the /api/export endpoint to create a workbook blob.
 *
 * @returns {Promise<Blob>} Promise that resolves to a workbook blob
 * @throws {Error} If workbook generation fails
 */
async function generateQueryWorkbook() {
    try {
        console.log('Starting query workbook generation...');
        
        // Get QueryInfo data from the active tab
        const activeTabData = getActiveTabData();
        const queryInfo = activeTabData.queryInfo;
        
        if (!queryInfo) {
            throw new Error('No query data available. Please ensure the PQQuery tab is active and has query information.');
        }
        
        console.log('QueryInfo collected:', queryInfo);
        
        // Get FileConfig data from configuration
        const fileConfigs = collectFileConfigData();
        console.log('FileConfig collected:', fileConfigs);
        
        // Prepare data for API call
        const requestData = {
            queryInfo: queryInfo,
            fileConfig: fileConfigs
        };
        
        // Call the server-side API endpoint
        console.log('Calling /api/export endpoint...');
        const response = await fetch('/api/export', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: 'Unknown server error' }));
            throw new Error(`Server error: ${errorData.message || response.statusText}`);
        }
        
        // Get the blob from the response
        const workbookBlob = await response.blob();
        
        console.log('Query workbook generated successfully via API');
        return workbookBlob;
        
    } catch (error) {
        console.error('Error generating query workbook:', error);
        
        // Log additional details for debugging
        if (error.message) {
            console.error('Error message:', error.message);
        }
        if (error.stack) {
            console.error('Error stack:', error.stack);
        }
        
        // Re-throw the error for caller to handle
        throw new Error(`Failed to generate query workbook: ${error.message || 'Unknown error'}`);
    }
}

/**
 * Download workbook blob using browser APIs
 * @param {Blob} blob - The workbook blob to download
 * @param {string} filename - The filename for the download
 */
async function downloadWorkbookBlob(blob, filename) {
    try {
        // Create download link
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        
        // Trigger download
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        // Clean up object URL
        URL.revokeObjectURL(url);
        
        console.log('Download triggered for:', filename);
    } catch (error) {
        console.error('Error downloading workbook:', error);
        throw error;
    }
}

/**
 * Open workbook in Excel Web using browser APIs
 * @param {Blob} blob - The workbook blob to open
 * @param {string} filename - The filename for the workbook
 */
async function openInExcelWeb(blob, filename) {
    try {
        console.log('Opening workbook in Excel Web via backend workbookManager...');
        
        // Create FormData to send blob to backend where workbookManager is available
        const formData = new FormData();
        formData.append('workbook', blob, filename);
        
        const response = await fetch('/api/excel-web/open', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const errorData = await response.json().catch(() => ({ message: 'Unknown error' }));
            throw new Error(`Backend error: ${errorData.message || response.statusText}`);
        }
        
  
    } catch (error) {
        console.error('Error opening workbook in Excel Web:', error);
        throw error;
    }
}