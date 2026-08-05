<p align="center">
  <img src="https://raw.githubusercontent.com/haris-musa/excel-mcp-server/main/assets/logo.png" alt="Excel MCP Server Logo" width="300"/>
</p>

[![PyPI version](https://img.shields.io/pypi/v/excel-mcp-server.svg)](https://pypi.org/project/excel-mcp-server/)
[![Total Downloads](https://static.pepy.tech/badge/excel-mcp-server)](https://pepy.tech/project/excel-mcp-server)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![smithery badge](https://smithery.ai/badge/@haris-musa/excel-mcp-server)](https://smithery.ai/server/@haris-musa/excel-mcp-server)
[![Install MCP Server](https://cursor.com/deeplink/mcp-install-dark.svg)](https://cursor.com/install-mcp?name=excel-mcp-server&config=eyJjb21tYW5kIjoidXZ4IGV4Y2VsLW1jcC1zZXJ2ZXIgc3RkaW8ifQ%3D%3D)

A Model Context Protocol (MCP) server that lets you manipulate Excel files without needing Microsoft Excel installed. Create, read, and modify Excel workbooks with your AI agent.

## Features

- 📊 **Excel Operations**: Create, read, update workbooks and worksheets
- 📈 **Data Manipulation**: Formulas, formatting, charts, pivot tables, and Excel tables
- 🔍 **Data Validation**: Built-in validation for ranges, formulas, and data integrity
- 🎨 **Formatting**: Font styling, colors, borders, alignment, and conditional formatting
- 📋 **Table Operations**: Create and manage Excel tables with custom styling
- 📊 **Chart Creation**: Generate various chart types (line, bar, pie, scatter, etc.)
- 🔄 **Pivot Tables**: Create dynamic pivot tables for data analysis
- 🔧 **Sheet Management**: Copy, rename, delete worksheets with ease
- 🔌 **Triple transport support**: stdio, SSE (deprecated), and streamable HTTP
- 🌐 **Remote & Local**: Works both locally and as a remote service

## Usage

The server supports three transport methods:

### 1. Stdio Transport (for local use)

```bash
uvx excel-mcp-server stdio
```

```json
{
   "mcpServers": {
      "excel": {
         "command": "uvx",
         "args": ["excel-mcp-server", "stdio"]
      }
   }
}
```

### 2. SSE Transport (Server-Sent Events - Deprecated)

```bash
uvx excel-mcp-server sse
```

**SSE transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/sse",
      }
   }
}
```

### 3. Streamable HTTP Transport (Recommended for remote connections)

```bash
uvx excel-mcp-server streamable-http
```

**Streamable HTTP transport connection**:
```json
{
   "mcpServers": {
      "excel": {
         "url": "http://localhost:8000/mcp",
      }
   }
}
```

## Environment Variables & File Path Handling

This server does **not** auto-load a `.env` file. Set variables in the process environment
(shell, Docker Compose, Kubernetes secrets, etc.) before starting the server.

### Core server variables

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `EXCEL_FILES_PATH` | SSE / streamable-http filepath tools | `./excel_files` | Directory used by local filepath-based tools |
| `FASTMCP_HOST` | No | `0.0.0.0` | Bind address for HTTP transports |
| `FASTMCP_PORT` | No | `8017` | Listen port for HTTP transports |

### SSE and Streamable HTTP Transports

When running the server with the **SSE or Streamable HTTP protocols**, you **must set the `EXCEL_FILES_PATH` environment variable on the server side**. This variable tells the server where to read and write Excel files.
- If not set, it defaults to `./excel_files`.
- With these transports, tool `filepath` values must be **relative** to that directory (e.g. `reports/q1.xlsx`); absolute paths and directory traversal are rejected.

You can also set the `FASTMCP_PORT` environment variable to control the port the server listens on (default is `8017` if not set).
- Example (Windows PowerShell):
  ```powershell
  $env:EXCEL_FILES_PATH="E:\MyExcelFiles"
  $env:FASTMCP_PORT="8007"
  uvx excel-mcp-server streamable-http
  ```
- Example (Linux/macOS):
  ```bash
  EXCEL_FILES_PATH=/path/to/excel_files FASTMCP_PORT=8007 uvx excel-mcp-server streamable-http
  ```

### Stdio Transport

When using the **stdio protocol**, the file path is provided with each tool call, so you do **not** need to set `EXCEL_FILES_PATH` on the server. The server will use the path sent by the client for each operation.

### Remote workbook job (`execute_workbook_job`)

Workflows integrations call the atomic tool `execute_workbook_job`. That tool does **not**
receive AWS credentials or raw file bytes. The client (workflows) signs short-lived S3 GET/PUT
URLs and passes them as arguments; this server downloads, mutates, and uploads against those URLs.

These variables are required on **this** Excel MCP process (local, Docker, or deployed), not in
the workflows app `.env`:

| Variable | Required | Default | Purpose |
|---|---|---|---|
| `EXCEL_MCP_ALLOWED_URL_HOSTS` | **Yes** (for remote jobs) | _(empty — fails closed)_ | Comma-separated **exact** hostnames allowed in signed URLs. Blocks SSRF by rejecting any other host. |
| `EXCEL_MCP_MAX_FILE_BYTES` | No | `104857600` (100 MB) | Max workbook size for download/upload |
| `EXCEL_MCP_REQUEST_TIMEOUT_SECONDS` | No | `120` | HTTP timeout for signed GET/PUT |

**`EXCEL_MCP_ALLOWED_URL_HOSTS` rules**

- Exact hostname match only (no wildcards, no path prefixes).
- Use the virtual-hosted S3 hostname: `<bucket>.s3.<region>.amazonaws.com`.
- Multiple hosts are comma-separated (e.g. local + staging buckets).
- If unset or empty, remote jobs fail with: `EXCEL_MCP_ALLOWED_URL_HOSTS is not configured`.

**Local development example** (PowerShell — same window you start the server from):

```powershell
$env:EXCEL_MCP_ALLOWED_URL_HOSTS = "mimasa-workflows-local-dev.s3.ap-south-1.amazonaws.com"
$env:EXCEL_MCP_MAX_FILE_BYTES = "104857600"
$env:EXCEL_MCP_REQUEST_TIMEOUT_SECONDS = "120"
$env:FASTMCP_PORT = "8017"
uv run excel-mcp-server streamable-http
```

**Docker / deployment example**

```yaml
environment:
  EXCEL_MCP_ALLOWED_URL_HOSTS: "your-bucket.s3.ap-south-1.amazonaws.com"
  EXCEL_MCP_MAX_FILE_BYTES: "104857600"
  EXCEL_MCP_REQUEST_TIMEOUT_SECONDS: "120"
  # FASTMCP_PORT is optional; defaults to 8017
  # FASTMCP_PORT: "8017"
```

Ensure the container port mapping and workflows `EXCEL_MCP_URL` match the listen port and path
(e.g. `http://excel-mcp:8017/mcp`).

These are **not** AWS access keys. Do not put `AWS_ACCESS_KEY_ID` / `AWS_SECRET_ACCESS_KEY` on
this service — workflows owns signing; this service only uses the signed URLs it receives.

## Available Tools

The server provides a comprehensive set of Excel manipulation tools. See [TOOLS.md](TOOLS.md) for complete documentation of all available tools.

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=haris-musa/excel-mcp-server&type=Date)](https://www.star-history.com/#haris-musa/excel-mcp-server&Date)

## License

MIT License - see [LICENSE](LICENSE) for details.
