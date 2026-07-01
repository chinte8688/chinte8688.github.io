# AGENTS.md

This is a document archive (正修科技大學 / Cheng Shiu University — 科技藝術發展處 / 藝文處). Not a software project.

## Structure

- Root: personal OneDrive folder with administrative documents, budgets, reports, teaching materials, and asset records.
- `mibau/`: empty; intended project directory.
- `cursor/`, `new_web_test/`, `wordpress-5.8.1-zh_TW/`, others: third-party or experimental subfolders.

## Commands

No build, test, lint, or typecheck commands exist. Python scripts (`.py`) in the root or under `Dropbox/` are standalone exercises (sorting, minesweeper, etc.) — run directly with `python <file>.py`.

## Conventions

- File names are in Traditional Chinese; paths may contain spaces and special characters.
- `.gitignore` does not exist — be careful not to commit large binaries (`.xlsx`, `.docx`, `.pdf`, `.pptx`, `.zip`, `.7z`, `.png`, `.jpg`, `.mp3`, `.kfx`), temp files (`~$*`, `~WRL*.tmp`), or SSH keys.
- No existing `AGENTS.md`, `CLAUDE.md`, `.cursorrules`, or `opencode.json` were found.

## Key directories

| Directory | Purpose |
|-----------|---------|
| `cursor/` | Experiment area (Python scripts) |
| `new_web_test/` | Web dev experimentation (MySQL, PHP, etc.) |
| `entryform(報名系統)/` | Registration system prototype |
| `wordpress-5.8.1-zh_TW/` | WordPress installation archive |

## Note

The git repo (root) has zero commits and no staged files. First meaningful commit should add a `.gitignore` excluding document binaries.
