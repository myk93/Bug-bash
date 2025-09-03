import { QueryInfo, FileConfig, ExportRequest, UserSessionState } from '../types';
import {workbookManager} from '@microsoft/connected-workbooks';
// Import Microsoft Connected Workbooks for frontend Excel Web integration


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
export const collectFileConfigData = (state: UserSessionState): FileConfig => {
  const fileConfig: FileConfig = {};
  
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
  
  // Create TempleteSettings object and filter out empty values
  const templeteSettings: Record<string, string> = {};
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
 * Download workbook blob using browser APIs
 */
export const downloadWorkbookBlob = async (blob: Blob, filename: string): Promise<void> => {
 workbookManager.downloadWorkbook(blob, filename);
};



/**
 * Open workbook in Excel Web using browser APIs
 * First tries to use Connected Workbooks library, falls back to backend API
 */
export const openInExcelWeb = async (blob: Blob, filename: string): Promise<void> => {
  workbookManager.openInExcelWeb(blob, filename);
};

/**
 * Handle Excel toggle functionality
 * Checks toggle state and calls appropriate function (Excel Web or Download)
 */
export const handleExcelToggle = async (
  blob: Blob, 
  isDownloadMode: boolean, 
  activeTab: string
): Promise<void> => {
  try {
    if (!blob) {
      throw new Error('No workbook blob provided');
    }
    
    // Generate filename
    const filename = generateFileName(activeTab);
    
    console.log('Excel toggle state:', isDownloadMode ? 'Download' : 'Excel Web');
    console.log('Processing with filename:', filename);
    
    if (isDownloadMode) {
      // Download mode - trigger browser download
      console.log('Triggering browser download...');
      await downloadWorkbookBlob(blob, filename);
      return;
    } else {
      // Excel Web mode - open in Excel Web
      console.log('Opening in Excel Web...');
      await openInExcelWeb(blob, filename);
      return;
    }
    
  } catch (error) {
    console.error('Error handling Excel toggle:', error);
    throw error;
  }
};