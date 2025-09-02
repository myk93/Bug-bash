const express = require('express');
const { v4: uuidv4 } = require('uuid');
const session = require('express-session');
const multer = require('multer');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');

const app = express();
const port = 3000;

// UserSession data model class
class UserSession {
    constructor() {
        this.sessionId = uuidv4();
        this.createdAt = new Date();
        this.lastActivity = new Date();
        this.uiState = {
            activeTab: 'grid',
            excelToggle: false
        };
        this.workspaceData = {
            gridData: [],
            tableData: [],
            pqQueryData: [],
            uploads: []
        };
    }

    updateLastActivity() {
        this.lastActivity = new Date();
    }

    updateUiState(newUiState) {
        this.uiState = { ...this.uiState, ...newUiState };
        this.updateLastActivity();
    }

    updateWorkspaceData(newWorkspaceData) {
        this.workspaceData = { ...this.workspaceData, ...newWorkspaceData };
        this.updateLastActivity();
    }

    resetSession() {
        this.uiState = {
            activeTab: 'grid',
            excelToggle: false
        };
        this.workspaceData = {
            gridData: [],
            tableData: [],
            pqQueryData: [],
            uploads: []
        };
        this.updateLastActivity();
    }

    toJSON() {
        return {
            sessionId: this.sessionId,
            createdAt: this.createdAt,
            lastActivity: this.lastActivity,
            uiState: this.uiState,
            workspaceData: this.workspaceData
        };
    }
}

// In-memory session storage
const sessions = new Map();

// Security middleware
app.use(helmet({
    contentSecurityPolicy: {
        directives: {
            defaultSrc: ["'self'"],
            styleSrc: ["'self'", "'unsafe-inline'"],
            scriptSrc: ["'self'", "'unsafe-inline'"],
            imgSrc: ["'self'", "data:", "https:"]
        }
    }
}));

app.use(cors({
    origin: process.env.NODE_ENV === 'production' ? false : true,
    credentials: true
}));

// Rate limiting
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 100, // limit each IP to 100 requests per windowMs
    message: 'Too many requests from this IP, please try again later.',
    standardHeaders: true,
    legacyHeaders: false
});

app.use('/api/', limiter);

// Session management middleware
app.use(session({
    secret: process.env.SESSION_SECRET || 'default-session-secret-change-in-production',
    resave: false,
    saveUninitialized: false,
    cookie: {
        secure: process.env.NODE_ENV === 'production',
        httpOnly: true,
        maxAge: 24 * 60 * 60 * 1000 // 24 hours
    }
}));

// Multer configuration for file uploads
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    },
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['text/csv', 'application/vnd.ms-excel', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
        if (allowedTypes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error('Invalid file type. Only CSV and Excel files are allowed.'), false);
        }
    }
});

// Body parsing middleware
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Static file serving (maintain backward compatibility)
app.use(express.static('.'));

// Input validation helpers
function validateSessionId(sessionId) {
    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
    return uuidRegex.test(sessionId);
}

function validateUiState(uiState) {
    if (!uiState || typeof uiState !== 'object') return false;
    
    if (uiState.activeTab && !['grid', 'table', 'pq-query'].includes(uiState.activeTab)) {
        return false;
    }
    
    if (uiState.excelToggle !== undefined && typeof uiState.excelToggle !== 'boolean') {
        return false;
    }
    
    return true;
}

function validateWorkspaceData(workspaceData) {
    if (!workspaceData || typeof workspaceData !== 'object') return false;
    
    const allowedKeys = ['gridData', 'tableData', 'pqQueryData', 'uploads'];
    const providedKeys = Object.keys(workspaceData);
    
    return providedKeys.every(key => allowedKeys.includes(key)) &&
           providedKeys.every(key => Array.isArray(workspaceData[key]));
}

// Session cleanup mechanism (runs every hour)
function cleanupSessions() {
    const now = new Date();
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours
    
    for (const [sessionId, session] of sessions.entries()) {
        if (now - session.lastActivity > maxAge) {
            sessions.delete(sessionId);
            console.log(`Cleaned up expired session: ${sessionId}`);
        }
    }
}

setInterval(cleanupSessions, 60 * 60 * 1000); // Run every hour

// API Routes

// POST /api/session/init - Create new session
app.post('/api/session/init', (req, res) => {
    try {
        const newSession = new UserSession();
        sessions.set(newSession.sessionId, newSession);
        
        console.log(`Created new session: ${newSession.sessionId}`);
        
        res.status(201).json({
            success: true,
            sessionId: newSession.sessionId,
            message: 'Session created successfully',
            session: newSession.toJSON()
        });
    } catch (error) {
        console.error('Error creating session:', error);
        res.status(500).json({
            success: false,
            message: 'Internal server error'
        });
    }
});

// GET /api/session/:sessionId - Get session data
app.get('/api/session/:sessionId', (req, res) => {
    try {
        const { sessionId } = req.params;
        
        if (!validateSessionId(sessionId)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid session ID format'
            });
        }
        
        const session = sessions.get(sessionId);
        if (!session) {
            return res.status(404).json({
                success: false,
                message: 'Session not found'
            });
        }
        
        session.updateLastActivity();
        
        res.json({
            success: true,
            session: session.toJSON()
        });
    } catch (error) {
        console.error('Error retrieving session:', error);
        res.status(500).json({
            success: false,
            message: 'Internal server error'
        });
    }
});

// PUT /api/session/:sessionId/state - Update session state
app.put('/api/session/:sessionId/state', (req, res) => {
    try {
        const { sessionId } = req.params;
        const { uiState, workspaceData } = req.body;
        
        if (!validateSessionId(sessionId)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid session ID format'
            });
        }
        
        const session = sessions.get(sessionId);
        if (!session) {
            return res.status(404).json({
                success: false,
                message: 'Session not found'
            });
        }
        
        // Validate and update UI state if provided
        if (uiState !== undefined) {
            if (!validateUiState(uiState)) {
                return res.status(400).json({
                    success: false,
                    message: 'Invalid UI state data'
                });
            }
            session.updateUiState(uiState);
        }
        
        // Validate and update workspace data if provided
        if (workspaceData !== undefined) {
            if (!validateWorkspaceData(workspaceData)) {
                return res.status(400).json({
                    success: false,
                    message: 'Invalid workspace data'
                });
            }
            session.updateWorkspaceData(workspaceData);
        }
        
        console.log(`Updated session state: ${sessionId}`);
        
        res.json({
            success: true,
            message: 'Session state updated successfully',
            session: session.toJSON()
        });
    } catch (error) {
        console.error('Error updating session state:', error);
        res.status(500).json({
            success: false,
            message: 'Internal server error'
        });
    }
});

// DELETE /api/session/:sessionId/reset - Reset session
app.delete('/api/session/:sessionId/reset', (req, res) => {
    try {
        const { sessionId } = req.params;
        
        if (!validateSessionId(sessionId)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid session ID format'
            });
        }
        
        const session = sessions.get(sessionId);
        if (!session) {
            return res.status(404).json({
                success: false,
                message: 'Session not found'
            });
        }
        
        session.resetSession();
        
        console.log(`Reset session: ${sessionId}`);
        
        res.json({
            success: true,
            message: 'Session reset successfully',
            session: session.toJSON()
        });
    } catch (error) {
        console.error('Error resetting session:', error);
        res.status(500).json({
            success: false,
            message: 'Internal server error'
        });
    }
});

// File upload endpoint (bonus feature for workspace data)
app.post('/api/session/:sessionId/upload', upload.single('file'), (req, res) => {
    try {
        const { sessionId } = req.params;
        
        if (!validateSessionId(sessionId)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid session ID format'
            });
        }
        
        const session = sessions.get(sessionId);
        if (!session) {
            return res.status(404).json({
                success: false,
                message: 'Session not found'
            });
        }
        
        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'No file uploaded'
            });
        }
        
        const uploadInfo = {
            originalName: req.file.originalname,
            mimetype: req.file.mimetype,
            size: req.file.size,
            uploadedAt: new Date(),
            data: req.file.buffer.toString('base64')
        };
        
        session.workspaceData.uploads.push(uploadInfo);
        session.updateLastActivity();
        
        console.log(`File uploaded to session: ${sessionId}`);
        
        res.json({
            success: true,
            message: 'File uploaded successfully',
            uploadInfo: {
                originalName: uploadInfo.originalName,
                mimetype: uploadInfo.mimetype,
                size: uploadInfo.size,
                uploadedAt: uploadInfo.uploadedAt
            }
        });
    } catch (error) {
        console.error('Error uploading file:', error);
        res.status(500).json({
            success: false,
            message: 'Internal server error'
        });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({
        success: true,
        message: 'Server is running',
        timestamp: new Date(),
        activeSessions: sessions.size
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                message: 'File too large. Maximum size is 10MB.'
            });
        }
    }
    
    console.error('Unhandled error:', error);
    res.status(500).json({
        success: false,
        message: 'Internal server error'
    });
});

// 404 handler for API routes
app.use('/api/*', (req, res) => {
    res.status(404).json({
        success: false,
        message: 'API endpoint not found'
    });
});

// Start server
app.listen(port, () => {
    console.log(`Server listening at http://localhost:${port}`);
    console.log(`Session management backend is ready for multi-user support`);
    console.log(`Active sessions will be cleaned up after 24 hours of inactivity`);
});

module.exports = app;