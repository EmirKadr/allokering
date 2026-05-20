# CLAUDE.md - AI Assistant Guide for `allokering`

Las `AGENTS.md` forst. Den filen innehaller repo-reglerna for hur nya funktioner ska byggas sa att GUI och CLI delar samma logik.

## Project Snapshot

- Main application: `allokering12.1.py`
- Current app metadata: `app_info.py`
- Current version: `12.1.4`
- Language: Swedish for user-facing text, reports, help texts, and most domain naming
- Runtime: Python desktop app with both GUI and CLI entry points

This repo is still centered around one large application file, but it is no longer "GUI only". The app now supports both:

- GUI for human users
- CLI for automation, testing, and agent-driven simulations

## Current Repository Structure

Key files and folders:

- `allokering12.1.py` - main GUI and CLI application
- `app_info.py` - app identity, version, GitHub release metadata, analytics config
- `update_service.py` - update check and installer download logic
- `analytics_store.py` - local analytics event storage
- `analytics_dashboard.py` - dashboard for reading local analytics files
- `wms_sok79.py` - WMS search support used by Eftersok
- `tests/` - pytest suite
- `TESTING.md` - how to run tests and CLI commands
- `AGENTS.md` - repo rules for agent-friendly implementation
- `../ALLOKERING_FILKUNSKAP.md` - shared file and column knowledge for CLI-facing flows across `projects`
- `packaging/windows/` - Windows installer and packaging assets

There is no database and no server component. The app processes CSV/XLSX inputs in memory and writes reports to text, CSV, or XLSX files.

## Running

GUI:

```powershell
python allokering12.1.py
```

CLI help:

```powershell
python allokering12.1.py --help
```

Run tests:

```powershell
pip install -r requirements-dev.txt
python -m pytest -q
```

## Current CLI Commands

The following commands exist today:

- `allocate`
- `ordersaldo`
- `lyx`
- `pafyllnadsprio`
- `hib-koppling`
- `overview-check`
- `dispatch-check`
- `vecka27-check`
- `eftersok`
- `prognos-report`
- `observations-update`
- `observations-sync`
- `split-values`
- `update-check`

When adding new report-like or data-processing features, prefer building them so they can join this list.

## Architecture Guidance

Use this mental model:

1. Input loading and normalization
2. Shared workflow or domain logic
3. CLI adapter
4. GUI adapter

Avoid pushing business logic deeper into `tkinter` handlers than necessary.

Preferred behavior:

- Shared workflow returns data, summaries, warnings, and report content
- CLI writes files and machine-readable summaries
- GUI shows messages, buttons, and "open in Excel" behavior

If GUI and CLI need the same feature, they should call the same underlying workflow.

## Testing Status

Pytest is present and should be extended when behavior changes.

Current test areas include:

- CLI end-to-end tests in `tests/cli/test_cli_commands.py`
- update logic tests in `tests/services/test_update_service.py`
- analytics storage tests in `tests/services/test_analytics_store.py`

When changing behavior, prefer adding:

- service-level tests for logic and normalization
- CLI end-to-end tests for full workflows

See `AGENTS.md` for the stronger "build it CLI-friendly" rule and a backlog of high-value next tests.

## Analytics and Updates

The repo now includes lightweight local analytics support:

- events are stored locally via `analytics_store.py`
- configuration defaults live in `app_info.py`
- the dashboard reads those local files via `analytics_dashboard.py`

Update handling is also built in:

- release metadata comes from GitHub releases
- update logic lives in `update_service.py`
- Windows packaging scripts live in `packaging/windows/`

## GitHub Workflows

Current workflows live in `.github/workflows/`:

- `auto-merge-claude.yml`
- `merge-observations.yml`
- `windows-release.yml`

Do not assume the repo is still using the old "single GUI file, no tests, no CLI" model. That is outdated.

## What Is Still True

- The main application logic is still concentrated in `allokering12.1.py`
- User-facing behavior should stay Swedish
- Input formats can vary, so flexible column matching and defensive normalization still matter
- Temporary exports and Excel-friendly outputs are still a central part of the workflow

## What Is No Longer True

The following old assumptions are wrong and should not be repeated:

- "There is no test suite"
- "There is no CLI"
- "There are no requirements files"
- "Everything lives in one file only"
- "The app has no analytics or update support"

If you are unsure how to extend the app, start with `AGENTS.md`, then `TESTING.md`, then inspect the closest existing CLI workflow in `allokering12.1.py`.
