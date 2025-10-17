# Planner Tasks — Deno App Structure

This document describes the architecture, key features, build/deploy tasks, and platform notes for the `deno-app/` project.

## Overview

- **Purpose**: Upload CSV/Excel, map columns, and create Microsoft Planner tasks via Microsoft Graph.
- **Stack**: Deno runtime, native HTTP (`Deno.serve`), no framework. Frontend is vanilla JS + Tailwind classes.
- **Executable targets**: macOS, Windows, Linux via `deno compile` with embedded static assets.

## Directory Layout

- `deno-app/`
  - `server.ts`: HTTP server, API routes, static serving, startup behaviors.
  - `auth.ts`: Authentication utilities for Microsoft Graph (interactive code flow).
  - `file-parser.ts`: CSV/Excel ingestion and normalization.
  - `static/`
    - `index.html`: Single-page UI.
    - `app.js`: Frontend logic.
  - `assets/icons/`: Platform icons (`logo.ico`, `logo.icns` if present).
  - `deno.json`: Tasks and compile settings (asset includes, targets, packaging).
- `documentation/`
  - `structure.md`: This document.
- `scripts/`
  - `package-macos.sh`: Builds `.app` bundle and `.dmg` for macOS.
- Repo root
  - `logo.png`, `logo.svg`: Branding assets (also embedded in builds).

## Key Features

- **Interactive Microsoft Auth** (`auth.ts`)
  - Launches default browser to Microsoft login.
  - Receives redirect on a temporary local HTTP port and exchanges code for access token.
  - Uses well-known public client IDs; tenant is `common`.
  - Scopes: `https://graph.microsoft.com/.default` (delegated permissions per public client app).
- **CSV/Excel Processing** (`file-parser.ts`)
  - Parses uploaded file, extracts rows and available columns.
  - Handles semicolons/commas variants and basic date normalization.
- **Planner Operations** (`auth.ts`)
  - Fetch planners/groups and buckets.
  - Create buckets (optional) and tasks; update descriptions; assign users.
  - Helper utilities to parse display names and look up users/groups.
- **Frontend UX** (`static/app.js`)
  - Multi-step UI: Auth → Upload → Map Columns → Planner → Results.
  - Column auto-detection and preview.
  - Live bucket resolution + optional bucket creation plan.
  - Assignee selection UI with dropdown, search, and multi-select checkboxes.

## Backend HTTP Routes (in `server.ts`)

- `GET /` → Serves `index.html` (embedded/static).
- `GET /static/*` → Serves frontend assets (embedded first; FS fallback).
- `GET /logo.png`, `GET /logo.svg`, `GET /favicon.ico` → Branding files (embedded/FS).
- `GET /api/auth/status` → Auth status for current session.
- `POST /api/auth/interactive` → Triggers interactive auth flow; opens system browser.
- `POST /api/auth/signout` → Clears session.
- `POST /api/parse-file` → Multipart upload; parses CSV/Excel.
- `POST /api/process-data` → Applies column mapping and returns prepared tasks preview.
- `GET /api/planners` → Lists user planners/groups.
- `GET /api/planners/:plannerId/buckets` → Lists buckets for a planner.
- `GET /api/planners/:plannerId/members?groupId=...` → Retrieves planner members (group + shared users).
- `POST /api/lookup-assignees` → Resolves assignees for tasks by name.
- `POST /api/lookup-buckets` → Reconciles task buckets against live plan data.
- `POST /api/tasks` (internal calls) → Creates tasks per prepared plan.

Notes:
- CORS headers set permissively for local use.
- All endpoints use an `X-Session-ID` header to correlate session (frontend stores in `localStorage`).

## Static Asset Strategy

- Compile embeds: `static/index.html`, `static/app.js`, `logo.png`, `logo.svg` (see `deno.json`).
- `server.ts` first tries to read assets from `import.meta.url` bundle; falls back to filesystem paths.
- This enables fully portable binaries without external files.

## Auth Flow Details (`auth.ts`)

1. `POST /api/auth/interactive` calls `GraphAuth.authenticateInteractive()`.
2. Picks a free localhost port and starts a temporary HTTP listener.
3. Constructs Microsoft authorize URL with:
   - `client_id`: one of the known public client IDs (Azure CLI, Graph Explorer, Graph PowerShell).
   - `response_type=code`, `redirect_uri=http://localhost:<port>`.
   - `scope=https://graph.microsoft.com/.default`.
   - `tenant=common`, `prompt=select_account`, `state=<uuid>`.
4. Opens default browser:
   - Windows: `rundll32 url.dll,FileProtocolHandler <url>` (fallback to `cmd /c start "" "<url>"`).
   - macOS: `open <url>`.
   - Linux: `xdg-open <url>`.
5. Temporary server captures the `code` query param; exchanges for `access_token`.
6. Session established; frontend persists `sessionId`.

## Frontend Highlights (`static/app.js`)

- **Assignee Dropdown**
  - `.assignee-toggle` opens/closes menu; internal clicks prevented from bubbling.
  - `.assignee-check` multi-select updates `assigneeOverrides[idx]` and summaries.
  - `.assignee-search` filters candidate/member lists.
- **Bucket Inline Selection**
  - Checkbox list of missing buckets with counts; selecting toggles creation plan.
  - Debounced re-render prevents losing assignee selections during quick changes.
- **Mapping & Preview**
  - Auto-picks likely columns; shows preview with status/bucket/assignee summaries.

## Build & Run (`deno.json` tasks)

- `deno task start` → Run server.
- `deno task dev` → Run server with watch.
- Compile (embeds assets):
  - `deno task compile` → Generic local binary.
  - `deno task compile-windows` → `planner-tasks.exe` (icon included).
  - `deno task compile-macos` → Intel macOS binary.
  - `deno task compile-macos-arm` → Apple Silicon macOS binary.
  - `deno task compile-linux` → Linux x64 binary.
- macOS DMG (Apple Silicon):
  - `deno task package-macos-arm` → Produces `.app` and `.dmg` under `deno-app/dist/`.

Notes:
- When running raw `deno compile` in zsh, always quote `--include=...` patterns.
- Our tasks use explicit includes to avoid glob pitfalls.

## Platform Notes

- **Windows**
  - Console window is expected (`deno compile` uses console subsystem).
  - Browser launch uses `rundll32` to preserve query strings; fallback is quoted `cmd /c start`.
- **macOS**
  - `deno compile` cannot embed an app icon directly; use `.app` bundle + `Info.plist` via `scripts/package-macos.sh`.
- **Linux**
  - Standard x64 target; ensure `xdg-open` available for browser launch.

## Environment & Config

- No external config file required for local use.
- Time synchronization and network access to `login.microsoftonline.com` and `graph.microsoft.com` required.
- If your tenant blocks `.default` for public clients, scopes may need adjustment (see Troubleshooting).

## Troubleshooting

- **Port In Use (8080)**: Terminate prior instance.
  - `lsof -iTCP:8080 -sTCP:LISTEN -n -P` → `kill -9 <PID>`.
- **Auth opens without scope on Windows**: Ensure latest build (rundll32 launcher). Quoted `start` fallback is present.
- **UI assets missing in binary**: Verify compile task includes `static/index.html` and `static/app.js`.
- **Assignee dropdown closes when clicking**: Fixed by stopping event propagation inside menu and debounced re-renders.
- **Gatekeeper on macOS**: Unsigned apps may require right-click → Open, or codesigning for distribution.

## Data Flow (High-Level)

```mermaid
description
flowchart LR
  A[Upload CSV/Excel] --> B[Parse /api/parse-file]
  B --> C[Map Columns /api/process-data]
  C --> D[Select Planner/Bucket]
  D --> E[Lookup Members/Buckets]
  E --> F[Preview & Overrides]
  F --> G[Create Tasks]
  G --> H[Graph API]
```

## Release Checklist

- **Embed assets**: Confirm compile tasks include required static files and logos.
- **Auth sanity**: Test interactive auth on target OS.
- **Planner ops**: Verify listing planners/buckets, creating tasks, assigning users.
- **Packaging**: For macOS distribution, run DMG packaging script; for Windows, include `.exe` and any README/notes.
