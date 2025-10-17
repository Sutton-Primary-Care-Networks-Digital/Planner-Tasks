# WARP.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Project Overview
Microsoft Planner Task Creator is a Streamlit web application that creates Microsoft Planner tasks from CSV/Excel files with assignee support. It uses Microsoft Graph API with OAuth2 authentication to integrate with Microsoft 365 services.

## Common Development Commands

### Running the Application
```bash
# Install dependencies
pip install -r requirements.txt

# Run the Streamlit application
streamlit run app.py

# Alternative: Use the Windows launcher (if on Windows)
./run_app.bat
```

### Development Setup
```bash
# Create virtual environment
python -m venv venv

# Activate virtual environment (macOS/Linux)
source venv/bin/activate

# Activate virtual environment (Windows)
venv\Scripts\activate.bat

# Install dependencies
pip install -r requirements.txt
```

### Code Quality
```bash
# Format code (if using black)
black *.py

# Type checking (if using mypy)
mypy *.py

# Lint code (if using flake8)
flake8 *.py
```

## Architecture Overview

### Core Components
The application follows a modular architecture with three main components:

1. **`app.py`** - Main Streamlit application
   - Handles the web UI and user interaction flow
   - Manages authentication state in Streamlit session
   - Orchestrates file parsing, assignee lookup, and task creation
   - Implements progressive workflow: Auth → Upload → Parse → Assignee Lookup → Planner Selection → Task Creation

2. **`graph_auth.py`** - Microsoft Graph API integration
   - Manages OAuth2 authentication using MSAL (Microsoft Authentication Library)
   - Implements multiple fallback client IDs for better organization compatibility
   - Handles user search and assignee lookup via Microsoft Graph
   - Creates tasks and assignments in Microsoft Planner
   - Includes robust error handling for common enterprise restrictions

3. **`file_parser.py`** - File processing and data validation
   - Parses CSV and Excel files using pandas
   - Provides column mapping interface for flexible file formats
   - Normalizes date formats to ISO 8601 for Microsoft Graph compatibility
   - Enriches tasks with user lookup results
   - Validates task data before creation

### Key Architectural Patterns

**Session State Management**: The application heavily uses Streamlit's session state to maintain user progress through the multi-step workflow. Key state variables:
- `access_token`: Microsoft Graph authentication token
- `processed_tasks`: Parsed task data from uploaded files
- `enriched_tasks`: Tasks with assignee lookup results
- File keys for cache management

**Error Handling Strategy**: The app implements graceful degradation for common enterprise scenarios:
- Multiple OAuth client IDs to bypass admin restrictions
- Fallback authentication methods
- User-friendly error messages with suggested solutions
- Automatic token refresh handling

**Assignee Resolution**: Complex multi-stage user lookup process:
1. Parse assignee names from various formats (e.g., "John Doe (COMPANY)")
2. Search Microsoft Graph using multiple strategies (exact match, partial match, fuzzy search)
3. Cache lookup results to avoid API rate limiting
4. Handle failed lookups gracefully while still creating tasks

## Development Guidelines

### Authentication Handling
- The app uses public client authentication (no client secrets)
- Multiple client IDs are attempted to work around organizational restrictions  
- Always handle 401/403 responses by clearing tokens and prompting re-authentication
- Consider personal Microsoft accounts for development/testing

### Microsoft Graph API Usage
- Follow proper ETag handling for task updates (required by Planner API)
- Use separate endpoints for task creation (`/tasks`) and description updates (`/tasks/{id}/details`)  
- Handle rate limiting and implement exponential backoff for production usage
- Filter out specific planner types (e.g., those containing "NHS.net") as needed

### File Processing
- Support both CSV and Excel formats via pandas
- Implement flexible column mapping to handle various file structures
- Normalize dates to ISO 8601 format with UTC timezone
- Validate required fields (at minimum, task title) before processing

### Streamlit Best Practices  
- Use `st.rerun()` instead of deprecated `st.experimental_rerun()`
- Implement progress bars for long-running operations
- Cache expensive operations in session state
- Provide clear user feedback for all operations
- Handle file uploads with proper validation and preview

### Error Messages
- Provide actionable error messages with specific solutions
- Differentiate between authentication, authorization, and API errors
- Include troubleshooting steps for common organizational restrictions
- Show progress and status updates during bulk operations

## Key Dependencies
- `streamlit==1.29.0` - Web application framework
- `msal==1.25.0` - Microsoft Authentication Library for OAuth2
- `requests==2.31.0` - HTTP client for Microsoft Graph API calls  
- `pandas==2.1.4` - Data processing for CSV/Excel files
- `openpyxl==3.1.2` - Excel file support
- `python-dateutil==2.8.2` - Flexible date parsing

## Microsoft Graph API Integration Notes
- Uses Microsoft Graph v1.0 endpoints for production stability
- Requires specific scopes: `https://graph.microsoft.com/.default`  
- Implements proper ETag handling for concurrency control
- Supports both personal and organizational Microsoft accounts
- Includes fallback methods for restricted enterprise environments