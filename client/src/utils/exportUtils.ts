import { TemplateSettings,QueryInfo, FileConfigs, Grid, GridConfig } from '@microsoft/connected-workbooks/dist/types';
import { UserSessionState } from '../types';
import { workbookManager } from '@microsoft/connected-workbooks';


/**
 * Generate filename for Excel export
 * Creates a filename in the format: {tabName}_{dateTime}.xlsx
 */
export const generateFileName = (activeTab: string): string => {
  try {
    // Map backend tab names to user-friendly names
    const tabNameMapping: Record<string, string> = {
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
};

/**
 * Get current PQ Query data from state
 */
export const getPQQueryData = (state: UserSessionState): QueryInfo => {
  return {
    queryMashup: state.pqQuery.queryMashup,
    refreshOnOpen: state.pqQuery.refreshOnOpen,
    queryName: state.pqQuery.queryName
  };
};

/**
 * Get active tab data based on session manager state
 * Returns structured data for the currently active tab
 */
export const getActiveTabData = (state: UserSessionState) => {
  const activeTab = state.activeTab;
  
  let queryInfo: QueryInfo | null = null;
  let gridData: any = null;
  
  console.log('Active tab data:', activeTab);
  
  // Collect data based on active tab
  if (activeTab === 'pq-query') {
    // For PQQuery tab, collect QueryInfo data
    queryInfo = getPQQueryData(state);
  } else if (activeTab === 'grid') {
    // For Grid tab, use gridData from state
    gridData = state.gridData || null;
  }
  
  // Return structured object with active tab and appropriate data
  return {
    activeTab: activeTab,
    queryInfo: queryInfo,
    gridData: gridData
  };
};

/**
 * Collect FileConfig data from the state
 * Maps data to the FileConfig interface structure required by @microsoft/connected-workbooks
 * Only includes properties that have actual user-entered data (not empty strings)
 */
export const collectFileConfigData = (state: UserSessionState): FileConfigs => {
  const fileConfig: FileConfigs = {};
  
  // Get document properties
  const docProps = state.docProps;
  
  // Filter out empty document properties
  const filteredDocProps: Record<string, string> = {};
  Object.keys(docProps).forEach(key => {
    const value = docProps[key as keyof typeof docProps];
    if (value && value.trim() !== '') {
      filteredDocProps[key] = value.trim();
    }
  });
  
  // Only add docProps if it has any non-empty values
  if (Object.keys(filteredDocProps).length > 0) {
    fileConfig.docProps = filteredDocProps;
  }
  
  // Get file configs
  const fileConfigs = state.fileConfigs;

  // Create TemplateSettings object and filter out empty values
  console.log("table name" + fileConfigs.tableName);
  console.log("sheet name" + fileConfigs.sheetName);
  let templateSettings: TemplateSettings = {};
  if (fileConfigs.tableName && fileConfigs.tableName.trim() !== '') {
    templateSettings.tableName = fileConfigs.tableName.trim();
  }
  if (fileConfigs.sheetName && fileConfigs.sheetName.trim() !== '') {
    templateSettings.sheetName = fileConfigs.sheetName.trim();
  }
  console.log('Collected TemplateSettings:', templateSettings);
  // Only add TemplateSettings if it has any non-empty values
  fileConfig.templateSettings = templateSettings;
console.log('Collected FileConfig:', fileConfig);
    // Set templateFile to the most recently uploaded file if available
  if (state.uploads && state.uploads.length > 0) {
    // Use the most recent upload (last item in array)
    fileConfig.templateFile = state.uploads[state.uploads.length - 1];
  }
  
  return fileConfig;
};

/**
 * Generate query workbook using the server-side API
 * This function collects data from the currently active tab and file configurations,
 * then calls the /api/export endpoint to create a workbook blob.
 */
export const generateQueryWorkbook = async (state: UserSessionState): Promise<Blob> => {
    // Get QueryInfo data from the active tab
    const activeTabData = getActiveTabData(state);
    const queryInfo = activeTabData.queryInfo;
    
    if (!queryInfo) {
      throw new Error('No query data available. Please ensure the PQQuery tab is active and has query information.');
    }
    
    
    // Get FileConfig data from configuration
    const fileConfigs = collectFileConfigData(state);
    
    // Get the blob from the response
    const workbookBlob = await workbookManager.generateSingleQueryWorkbook(queryInfo,undefined,fileConfigs);
    
    return workbookBlob;
};

/**
 * Generate grid workbook using the workbook manager
 * This function creates a Grid object from the session state and generates a workbook blob.
 * It handles both HTML table mode and grid mode based on the toggle state.
 */
export const generateGridWorkbook = async (state: UserSessionState): Promise<Blob> => {
  try {
    // Check if we have grid data
    if (!state.gridData || !state.gridData.data || state.gridData.data.length === 0) {
      throw new Error('No grid data available. Please ensure the Grid tab has data.');
    }

    // Get FileConfig data from configuration
    const fileConfigs = collectFileConfigData(state);
    
    // Check the grid view toggle state
    const isGridView = state.gridView?.isGridView || false;
    
    if (isGridView) {
      // Grid mode - create Grid object and use generateTableWorkbookFromGrid
      console.log('Exporting in Grid mode...');
      
      // Convert string[][] to (string | number | boolean)[][]
      const gridData: (string | number | boolean)[][] = state.gridData.data.map(row =>
        row.map(cell => {
          // Try to convert to number if possible
          if (cell.trim() !== '' && !isNaN(Number(cell))) {
            return Number(cell);
          }
          // Try to convert to boolean if possible
          if (cell.toLowerCase() === 'true') return true;
          if (cell.toLowerCase() === 'false') return false;
          // Return as string
          return cell;
        })
      );
      
      // Create GridConfig from state
      const gridConfig: GridConfig = {
        promoteHeaders: state.gridView?.promoteHeaders || false,
        adjustColumnNames: state.gridView?.adjustColumnNames || false
      };
      
      // Create Grid object
      const grid: Grid = {
        data: gridData,
        config: gridConfig
      };
      
      console.log('Generating workbook from Grid object:', grid);
      const workbookBlob = await workbookManager.generateTableWorkbookFromGrid(grid, fileConfigs);
      return workbookBlob;
      
    } else {
      // HTML Table mode - get HTML table element and use generateTableWorkbookFromHtml
      console.log('Exporting in HTML Table mode...');
      
      // Find the HTML table element in the Grid tab
      const htmlTable = document.querySelector('#Grid .editable-grid') as HTMLTableElement;
      
      if (!htmlTable) {
        throw new Error('HTML table element not found. Please ensure the Grid tab is visible.');
      }
      
      console.log('Generating workbook from HTML table element:', htmlTable);
      const workbookBlob = await workbookManager.generateTableWorkbookFromHtml(htmlTable, fileConfigs);
      return workbookBlob;
    }
    
  } catch (error) {
    console.error('Error generating grid workbook:', error);
    throw error;
  }
};

/**
 * Download workbook blob using browser APIs
 */
export const downloadWorkbookBlob = async (blob: Blob, filename: string): Promise<void> => {
 workbookManager.downloadWorkbook(blob, filename);
};



/**
 * Open workbook in Excel Web using browser APIs
 * First tries to use Connected Workbooks library, falls back to backend API
 */
export const openInExcelWeb = async (file: Blob, filename?: string, allowTyping?: boolean, allowEdit?: boolean): Promise<void> => {
  workbookManager.openInExcelWeb(file, filename, allowTyping, allowEdit);
};

/**
 * Handle Excel toggle functionality
 * Checks toggle state and calls appropriate function (Excel Web or Download)
 */
export const handleExcelToggle = async (
  blob: Blob,
  isDownloadMode: boolean,
  activeTab: string,
  editMode: boolean = true
): Promise<void> => {
  try {
    if (!blob) {
      throw new Error('No workbook blob provided');
    }

    // Generate filename
    const filename = generateFileName(activeTab);

    console.log('Excel toggle state:', isDownloadMode ? 'Download' : 'Excel Web');
    console.log('Edit mode:', editMode ? 'Edit' : 'View');
    console.log('Processing with filename:', filename);

    if (isDownloadMode) {
      // Download mode - trigger browser download
      console.log('Triggering browser download...');
      await downloadWorkbookBlob(blob, filename);
      return;
    } else {
      // Excel Web mode - open in Excel Web with edit/view mode
      console.log('Opening in Excel Web...');
      const allowTyping = editMode;
      const allowEdit = editMode;
      await openInExcelWeb(blob, filename, allowTyping, allowEdit);
      return;
    }

  } catch (error) {
    console.error('Error handling Excel toggle:', error);
    throw error;
  }
};