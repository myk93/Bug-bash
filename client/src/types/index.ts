// Session Management Types
export interface UserSessionState {
  sessionId: string;
  activeTab: 'grid' | 'table' | 'pq-query';
  excelToggle: boolean;
  excelWebEditMode: boolean;
  sidebarCollapsed: boolean;
  fileConfigs: {
    tableName: string;
    sheetName: string;
  };
  docProps: {
    title: string;
    subject: string;
    keywords: string;
    createdBy: string;
    description: string;
    lastModifiedBy: string;
    category: string;
    revision: string;
  };
  pqQuery: {
    queryMashup: string;
    refreshOnOpen: boolean;
    queryName: string;
  };
  gridView: {
    isGridView: boolean;
    promoteHeaders: boolean;
    adjustColumnNames: boolean;
  };
  gridData: {
    data: string[][];
    rows: number;
    cols: number;
  };
  uploads: any[];
}

// SessionResponse interface removed - no longer needed for pure frontend app

// Query Types
export interface QueryInfo {
  queryMashup: string;
  refreshOnOpen: boolean;
  queryName?: string;
}

// File Config Types
export interface FileConfig {
  docProps?: {
    title?: string;
    subject?: string;
    keywords?: string;
    createdBy?: string;
    description?: string;
    lastModifiedBy?: string;
    category?: string;
    revision?: string;
  };
  TempleteSettings?: {
    tableName?: string;
    sheetName?: string;
  };
  templateFile?: any;
}

// Notification Types
export interface NotificationOptions {
  message: string;
  type?: 'info' | 'success' | 'error' | 'warning';
  duration?: number;
}

// Tab Types
export type TabName = 'Grid' | 'PQQuery';

// Export Request Types removed - no longer needed for pure frontend app