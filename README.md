# React Data Grid Application

A pure frontend React application with data grid functionality, Power Query integration, and Microsoft Connected Workbooks support.

## Overview

This application has been converted from a full-stack Express.js + vanilla JavaScript application to a pure React frontend application. All state management is handled locally using localStorage, and no backend server is required.

## Features

- **Interactive Data Grid**: Editable 3x3 grid with cell navigation and resize functionality
- **Power Query Editor**: M language editor interface with example code buttons
- **File Upload**: Drag-and-drop file upload functionality (XLSX/CSV support)
- **Export Functionality**: Export to Excel Web or download using Microsoft Connected Workbooks
- **Session Management**: Local state persistence using localStorage
- **Document Properties**: Configurable document metadata
- **Responsive Design**: Clean, modern UI with sidebar and tabbed interface

## Tech Stack

- **Frontend**: React 18 with TypeScript
- **State Management**: Custom hooks with localStorage persistence
- **Styling**: CSS modules with responsive design
- **Export Library**: @microsoft/connected-workbooks
- **Build Tool**: Create React App

## Getting Started

### Prerequisites

- Node.js (version 14 or higher)
- npm or yarn package manager

### Installation

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd <project-directory>
   ```

2. **Install dependencies**
   ```bash
   cd client
   npm install
   ```

3. **Build the application**
   ```bash
   npm run build
   ```

### Running the Application

#### Option 1: Using serve (Recommended)
```bash
cd client
npx serve -s build -p 3000
```

Then open your browser and navigate to `http://localhost:3000`

#### Option 2: Using any static file server
You can serve the `client/build` directory using any static file server:

```bash
# Using Python 3
cd client/build
python -m http.server 3000

# Using PHP
cd client/build
php -S localhost:3000

# Using Live Server extension in VS Code
# Right-click on client/build/index.html and select "Open with Live Server"
```

### Development Mode

To run in development mode with hot reload:

```bash
cd client
npm start
```

This will start the development server at `http://localhost:3000`

## Project Structure

```
├── client/
│   ├── public/           # Static assets
│   ├── src/
│   │   ├── components/   # React components
│   │   │   ├── ExcelToggle.tsx
│   │   │   ├── GridTab.tsx
│   │   │   ├── Notification.tsx
│   │   │   ├── PQQueryTab.tsx
│   │   │   └── Sidebar.tsx
│   │   ├── hooks/        # Custom React hooks
│   │   │   ├── useNotification.ts
│   │   │   └── useSessionManager.ts
│   │   ├── types/        # TypeScript type definitions
│   │   │   └── index.ts
│   │   ├── utils/        # Utility functions
│   │   │   └── exportUtils.ts
│   │   ├── App.tsx       # Main application component
│   │   ├── App.css       # Application styles
│   │   └── index.tsx     # Application entry point
│   ├── build/            # Production build output
│   └── package.json      # Dependencies and scripts
├── app.js                # Legacy Express server (no longer needed)
├── index.html            # Legacy HTML file (no longer needed)
├── style.css             # Legacy styles (no longer needed)
├── script.js             # Legacy JavaScript (no longer needed)
└── README.md             # This file
```

## Application Components

### Main Features

1. **Grid Tab**: Interactive data grid with:
   - Editable cells with keyboard navigation
   - Grid view options (promote headers, adjust column names)
   - CSV import functionality

2. **PQ Query Tab**: Power Query editor with:
   - M language syntax support
   - Example code insertion
   - Query configuration options

3. **Sidebar**: Configuration panel with:
   - File upload (drag & drop)
   - Document properties
   - Export settings

4. **Excel Toggle**: Export functionality with:
   - Toggle between Excel Web and Download modes
   - Integration with Microsoft Connected Workbooks
   - Session reset capability

### State Management

- All state is managed locally using React hooks and localStorage
- Session data persists between browser sessions
- No server-side dependencies required

### Export Functionality

- Uses `@microsoft/connected-workbooks` library for Excel integration
- Supports both Excel Web opening and file download
- Fallback to CSV download if Excel integration fails

## Browser Compatibility

- Modern browsers with ES6+ support
- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## Troubleshooting

### Common Issues

1. **Application won't start**: Ensure all dependencies are installed with `npm install`
2. **Export not working**: Check that `@microsoft/connected-workbooks` is properly installed
3. **Local storage issues**: Clear browser cache and localStorage if experiencing state issues

### Development Issues

1. **TypeScript errors**: Run `npm run build` to check for compilation errors
2. **Missing dependencies**: Check `package.json` and run `npm install`
3. **Port conflicts**: Change the port number in the serve command

## License

This project is available under the MIT License.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

---

**Note**: This application was successfully converted from a full-stack application to a pure frontend React application. All server-side functionality has been replaced with client-side equivalents using localStorage and browser APIs.