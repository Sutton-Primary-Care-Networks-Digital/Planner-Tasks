# Microsoft Planner Task Creator (Deno)

A portable, desktop-friendly web app for creating Microsoft Planner tasks from CSV/Excel files. Built with Deno, TypeScript, and Tailwind CSS. Optimized for NHS.net environments and runs locally with no extra services.

## Features

- **Microsoft Authentication**: Browser-based OAuth2 with multiple client ID fallback
- **File Processing**: Support for CSV and Excel files with flexible column mapping
- **Assignee Management**: Automatic user lookup with multiple name format support
- **Bucket Management**: Smart bucket matching and creation with inline bucket selection
- **Portable Deployment**: Compiles to standalone executables for easy distribution
- **Modern UI**: Clean, responsive interface with Tailwind CSS

## Quick Start

### Prerequisites

- [Deno](https://deno.land/) v1.37 or later

### Development

1. Clone and navigate to the project:
```bash
cd deno-app
```

2. Start the development server:
```bash
deno task dev
```

3. Open your browser to `http://localhost:8080`

### Production Build

Compile to standalone executables for different platforms:

```bash
# Current platform
deno task compile

# Windows
deno task compile-windows

# macOS
deno task compile-macos

# Linux
deno task compile-linux
```

## Usage

1. **Authentication**: Click “Sign In”. A browser window opens for Microsoft login. On success, you return to the app.

2. **File Upload**: Upload a CSV or Excel file containing your task data.

3. **Column Mapping**: Map your file columns to Planner fields:
   - Title (required)
   - Description, Start Date, Due Date (optional)
   - Assignee (supports multiple formats)
   - Bucket Name, Status (optional)

4. **Planner Selection**: Choose your target planner and a default bucket.

5. **Task Preview & Buckets**:
   - Inline bucket selection lists all missing buckets (checkboxes on left). “Create all missing buckets” is on by default and acts as select-all.
   - Tasks whose bucket is missing and not selected will go to the default bucket.
   - Preview shows assignees (or “Not found”) and lets you manually pick members per task.

6. **Create**:
   - The app creates the selected missing buckets first, refreshes buckets, then creates tasks with assignees/status.

## File Format Support

### Supported Formats
- CSV (.csv)
- Excel (.xlsx, .xls)

### Column Examples
```csv
Title,Description,Due Date,Assignee,Bucket Name,Status
"Setup meeting","Plan project kickoff","2024-01-15","John Doe (COMPANY)","Planning","Not Started"
"Review design","Check UI mockups","2024-01-20","Jane Smith, Bob Wilson","Design","In Progress"
```

### Assignee Formats
- `John Doe (COMPANY)`
- `John Doe`
- `john.doe@company.com`
- Multiple assignees: `John Doe, Jane Smith`

## Authentication

The application uses Microsoft Graph API with OAuth2 authentication. It supports multiple public client IDs for better organizational compatibility:

- Microsoft Azure CLI
- Microsoft Graph Explorer  
- Microsoft Graph PowerShell

No additional Azure app registration required (public client authentication).

## Architecture

### Backend (Deno native HTTP)
- **`server.ts`**: Web server and API endpoints (native Deno HTTP, no Oak)
- **`auth.ts`**: Microsoft Graph authentication
- **`file-parser.ts`**: CSV/Excel processing

### Frontend (Vanilla JS + Tailwind)
- **`static/index.html`**: Main UI
- **`static/app.js`**: Client-side logic
- Progressive web app workflow with step-by-step guidance

## API Endpoints

- `GET /api/auth/status` - Check session auth status
- `POST /api/auth/interactive` - Interactive auth (opens browser)
- `GET /api/planners` - Get user's planners
- `GET /api/planners/:id/buckets` - Get planner buckets
- `GET /api/planners/:id/members` - Get members for a plan (and group)
- `POST /api/parse-file` - Parse uploaded file
- `POST /api/process-data` - Process column mapping
- `POST /api/lookup-assignees` - Lookup user assignees
- `POST /api/lookup-buckets` - Resolve tasks to existing buckets
- `POST /api/create-tasks` - Create tasks in Planner
- `POST /api/buckets/create` - Create one or more new buckets
- `POST /api/auth/signout` - Clear session

## Configuration

### Environment Variables
- `PORT`: Server port (default: 8080)

### Customization
- Client IDs can be modified in `auth.ts`
- UI styling can be customized via Tailwind classes
- File processing logic in `file-parser.ts`

## Deployment Options

### Standalone Executable
```bash
deno task compile
./planner-tasks
```

### Docker (Optional)
```dockerfile
FROM denoland/deno:alpine
WORKDIR /app
COPY . .
RUN deno cache main.ts
EXPOSE 8080
CMD ["deno", "run", "--allow-all", "main.ts"]
```

### Cloud Deployment
- Deploy to any platform supporting Deno (Deno Deploy, Railway, etc.)
- No database required - uses in-memory sessions
- Stateless design for easy scaling

## Troubleshooting

### Authentication Issues
- Ensure you're using a valid NHS.net account
- Try using incognito mode or different browser
- Verify access to Microsoft Planner within NHS.net
- Clear browser cache and cookies
- Expired session: the app shows a toast and automatically returns to the Sign In step.

### File Processing Issues
- Ensure CSV uses UTF-8 encoding
- Check for empty rows or columns
- Verify column headers are in first row
- Maximum file size depends on available memory

### NHS.net Environment
- The app includes multiple client ID fallbacks for NHS.net compatibility
- Browser-based authentication is optimized for NHS.net environments
- All planners within your NHS.net tenant will be accessible

## Development

### Project Structure
```
deno-app/
├── server.ts            # Main server (native Deno HTTP)
├── auth.ts              # Authentication module
├── file-parser.ts       # File processing
├── deno.json           # Deno configuration
├── static/
│   ├── index.html      # Main UI
│   └── app.js          # Frontend logic
└── README.md
```

### Adding Features
1. Backend: Add new endpoints in `server.ts`
2. Frontend: Extend `app.js` class methods
3. Authentication: Modify `auth.ts` for new Graph APIs
4. File Processing: Enhance `file-parser.ts` for new formats

## Security

- Uses Microsoft's public client authentication
- No client secrets stored
- Session-based state management
- CORS enabled for development
- Input validation on all endpoints

## Performance

- Minimal dependencies (Deno standard library + Oak)
- In-memory session storage (suitable for single-user deployment)
- Async/await for non-blocking operations
- Efficient CSV parsing with streaming support

## License

This project maintains the same license as the original Python version.

## Migration from Python Version

Key improvements in the Deno version:
- Single executable deployment
- No Python/pip dependencies
- Better cross-platform compatibility
- Modern web UI with progressive workflow
- Maintained feature parity with original functionality