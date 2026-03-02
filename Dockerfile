FROM python:3.11-slim

# Cache bust — forces full layer rebuild: 2026-03-02T21:52:00Z
ARG CACHE_BUST=2026-03-02T21:52:00Z

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy everything needed for install
COPY pyproject.toml README.md ./
COPY src/ ./src/

# Install package (non-editable so module is fully installed into site-packages)
RUN pip install --no-cache-dir .

# Expose port
EXPOSE 8080

# Start server
CMD ["uvicorn", "m365_mcp.server:app", "--host", "0.0.0.0", "--port", "8080"]
