# Microsoft Planner Task Creator

Create Microsoft Planner tasks from CSV/Excel files. This repo hosts multiple implementations; the **Deno app is the current primary**.

## Quick Start (Deno app)

The Deno version lives in `deno-app/` and runs locally with a modern web UI.

```bash
cd deno-app
deno task dev
# open http://localhost:8080
```

See full docs in `deno-app/README.md`.

### Highlights
- Browser-based Microsoft sign-in (interactive flow)
- CSV/Excel upload with column mapping
- Inline bucket selection, with “Create all missing buckets” toggle
- Assignee lookup; manual selection per task
- Creates selected buckets then tasks via Microsoft Graph

## Other implementations

- **Python (original)**: `python-version/`
- **C# (.NET MAUI, experimental)**: `PlannerTaskCreator/` and `PlannerTaskCreatorConsole/`

## Repository structure

```
deno-app/               # Primary app (Deno)
python-version/         # Original Streamlit prototype
PlannerTaskCreator/     # .NET MAUI UI (experimental)
PlannerTaskCreatorConsole/  # .NET console (auth test)
README.md               # This file
```

## Licensing & security

- Uses Microsoft public client auth (no secrets)
- Session stored locally in-memory

## Support

For issues or feature requests, open a GitHub issue. For Deno usage details, refer to `deno-app/README.md`.