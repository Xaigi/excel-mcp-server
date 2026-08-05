# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

Excel MCP Server — a Model Context Protocol server that manipulates `.xlsx` files via `openpyxl` without requiring Microsoft Excel. Built on `FastMCP`. Distributed on PyPI as `excel-mcp-server` and packaged as `.mcpb` for Claude Desktop.

## Commands

Uses `uv` for env/dependency management (see `uv.lock`, `pyproject.toml`).

- Install deps: `uv sync`
- Run server locally (three transports):
  - `uv run excel-mcp-server stdio`
  - `uv run excel-mcp-server streamable-http` (recommended for remote)
  - `uv run excel-mcp-server sse` (deprecated)
- Run tests: `uv run python -m unittest discover -s tests -v`
- Run a single test: `uv run python -m unittest tests.test_sandbox_paths.TestGetExcelPathSandbox.test_stdio_accepts_absolute_only`

Relevant env vars for HTTP/SSE modes: `EXCEL_FILES_PATH` (default `./excel_files`), `FASTMCP_PORT` (default `8017`), `FASTMCP_HOST` (default `0.0.0.0`).

## Architecture

Two-layer design; keep the split intact when adding features:

1. **Implementation modules** in `src/excel_mcp/` — pure logic on top of `openpyxl`, each focused on one concern: `workbook.py`, `sheet.py`, `data.py`, `formatting.py`, `calculations.py`, `chart.py`, `pivot.py`, `tables.py`, `validation.py`, `cell_validation.py`, `cell_utils.py`. These raise typed errors from `exceptions.py` (`ValidationError`, `WorkbookError`, `SheetError`, `DataError`, `FormattingError`, `CalculationError`, `PivotError`, `ChartError`) and return dicts (often `{"message": ...}`).
2. **MCP tool layer** in `server.py` — thin `@mcp.tool(...)` wrappers with `ToolAnnotations` (title, readOnly/destructive hints) that call the impl modules (imported with `_impl` suffix), catch the typed exceptions, and return user-facing strings. Unknown exceptions are logged and re-raised. New tools go here and follow this same wrap/catch pattern.

`__main__.py` exposes the `stdio` / `sse` / `streamable-http` subcommands via Typer, each calling a `run_*` entry point in `server.py`. The `run_*` functions set `EXCEL_FILES_PATH` (except stdio) and start `FastMCP`.

### Path handling (security-critical)

All tool handlers route filepaths through `server.get_excel_path()`. Its rules differ by transport and must be preserved:
- **stdio** (`EXCEL_FILES_PATH is None`): only **absolute** paths accepted; relative paths raise `ValueError`.
- **SSE / streamable-http** (`EXCEL_FILES_PATH` set): only **relative** paths accepted, resolved under `EXCEL_FILES_PATH` via `os.path.realpath` and checked with `_resolved_path_is_within` to reject directory traversal and symlink escapes. Absolute paths raise `ValueError`.

This sandboxing is what `tests/test_sandbox_paths.py` verifies — any change to `get_excel_path` should keep those tests green.

### Logging

Logs go to `excel-mcp.log` at the repo root (path computed from `__file__` so stdio clients with unpredictable cwd still work). **Do not** add stdout logging — stdio transport requires stdout to carry only MCP protocol frames.

## Tool reference

`TOOLS.md` documents every exposed MCP tool (arguments and behavior). Update it when adding, removing, or changing tool signatures in `server.py`.

## Packaging

`manifest.json` + `excel-mcp-server-*.mcpb` are the Claude Desktop bundle. Bump `version` in both `pyproject.toml` and `manifest.json` together on release.
