FROM python:3.12-slim

RUN pip install --no-cache-dir uv==0.5.11

WORKDIR /app

COPY pyproject.toml uv.lock README.md ./
COPY src ./src

RUN uv sync --frozen --no-dev

ENV PATH="/app/.venv/bin:$PATH" \
    EXCEL_FILES_PATH=/data/excel_files \
    FASTMCP_HOST=0.0.0.0 \
    FASTMCP_PORT=8017

RUN mkdir -p /data/excel_files

EXPOSE 8017

CMD ["excel-mcp-server", "streamable-http"]
