---
description: Repository Information Overview
alwaysApply: true
---

# Microsoft Planner Task Creator Information

## Repository Summary

A tool for creating Microsoft Planner tasks from CSV/Excel files with enhanced functionality including multiple assignee formats, bucket auto-detection, and improved user experience. The repository contains multiple implementations of the same concept.

## Repository Structure

- **python-version/**: Original implementation using Streamlit web application
- **neutralino-app/**: Desktop application implementation using Neutralino.js
- **C# Version**: Referenced in README but not present in the repository yet

## Projects

### Python Version (Streamlit)

**Configuration File**: python-version/requirements.txt

#### Language & Runtime

**Language**: Python
**Version**: Compatible with Python 3.x
**Package Manager**: pip

#### Dependencies

**Main Dependencies**:

- streamlit==1.29.0
- msal==1.25.0
- requests==2.31.0
- pandas==2.1.4
- openpyxl==3.1.2
- python-dateutil==2.8.2

#### Build & Installation

```bash
cd python-version
pip install -r requirements.txt
streamlit run app.py
```

#### Main Files

**Entry Point**: python-version/app.py
**Key Components**:

- graph_auth.py: Microsoft Graph authentication
- file_parser.py: CSV/Excel file parsing

### Neutralino Desktop App

**Configuration File**: neutralino-app/neutralino.config.json

#### Language & Runtime

**Language**: JavaScript/HTML/CSS
**Version**: Neutralino.js 6.3.0
**Build System**: Neutralino CLI
**Package Manager**: npm

#### Dependencies

**Development Dependencies**:

- @neutralinojs/neu==10.2.0

#### Build & Installation

```bash
cd neutralino-app
npm install
npm run build
```

#### Main Files

**Entry Point**: neutralino-app/resources/index.html
**Configuration**: neutralino-app/neutralino.config.json

### C# Version (Planned)

According to the README, a C# version is in development but not yet present in the repository. It will include:

#### Planned Features

- .NET MAUI cross-platform desktop application
- Enhanced CSV processing with multiple assignee support
- Bucket auto-detection with fuzzy matching
- Modern MVVM architecture

#### Planned Dependencies

- Microsoft.Graph 5.36.0
- Microsoft.Identity.Client 4.56.0
- CsvHelper 30.0.1
- CommunityToolkit.Mvvm 8.2.2
