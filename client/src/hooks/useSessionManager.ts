import { useState, useEffect, useCallback } from 'react';
import { UserSessionState } from '../types';

const SESSION_STORAGE_KEY = 'user_session_state';

interface SessionManager {
  state: UserSessionState;
  updateState: (updates: Partial<UserSessionState>) => void;
  resetSession: () => void;
  isLoading: boolean;
  error: string | null;
}

const defaultState: UserSessionState = {
  sessionId: `local_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
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
    isGridView: false,
    promoteHeaders: false,
    adjustColumnNames: false
  },
  gridData: {
    data: [
      ['', '', ''],
      ['', '', ''],
      ['', '', '']
    ],
    rows: 3,
    cols: 3
  },
  uploads: []
};

const useSessionManager = (): SessionManager => {
  const [state, setState] = useState<UserSessionState>(defaultState);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  // Initialize session from localStorage
  useEffect(() => {
    const initializeSession = () => {
      try {
        setIsLoading(true);
        setError(null);

        // Try to load from localStorage
        const savedState = localStorage.getItem(SESSION_STORAGE_KEY);
        if (savedState) {
          const parsedState = JSON.parse(savedState);
          setState(parsedState);
          console.log('Loaded session from localStorage');
        } else {
          // Use default state and save it
          setState(defaultState);
          localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(defaultState));
          console.log('Created new local session');
        }

      } catch (error) {
        console.error('Failed to initialize session:', error);
        setError(`Failed to initialize: ${error instanceof Error ? error.message : 'Unknown error'}`);
        
        // Fallback to default state
        setState(defaultState);
        try {
          localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(defaultState));
        } catch (storageError) {
          console.error('Failed to save default state:', storageError);
        }
      } finally {
        setIsLoading(false);
      }
    };

    initializeSession();
  }, []);

  const updateState = useCallback((updates: Partial<UserSessionState>) => {
    setState(prevState => {
      const newState = { ...prevState, ...updates };
      
      // Save to localStorage immediately
      try {
        localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(newState));
        console.log('State updated and saved to localStorage');
      } catch (error) {
        console.error('Failed to save state to localStorage:', error);
        setError(`Failed to save: ${error instanceof Error ? error.message : 'Unknown error'}`);
      }
      
      return newState;
    });
  }, []);

  const resetSession = useCallback(() => {
    try {
      setError(null);
      
      // Create new default state with new session ID
      const newState = {
        ...defaultState,
        sessionId: `local_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
      };
      
      // Update state and localStorage
      setState(newState);
      localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(newState));
      
      console.log('Session reset successfully');

    } catch (error) {
      console.error('Error resetting session:', error);
      setError(`Reset failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      
      // Still reset locally
      setState(defaultState);
      try {
        localStorage.setItem(SESSION_STORAGE_KEY, JSON.stringify(defaultState));
      } catch (storageError) {
        console.error('Failed to save reset state:', storageError);
      }
    }
  }, []);

  return {
    state,
    updateState,
    resetSession,
    isLoading,
    error
  };
};

export default useSessionManager;