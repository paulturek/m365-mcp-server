FROM python:3.11-slim

# Cache bust — forces full layer rebuild: 2026-03-02T21:50:00Z
ARG CACHE_BUST=2026-03-02T21:50:00Z

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency files first for layer caching
COPY pyproject.toml ./

# Install Python dependencies
RUN pip install --no-cache-dir -e ".[prod]" 2>/dev/null || pip install --no-cache-dir -e .

# Copy source code (after deps to maximise cache hits on clean builds)
COPY src/ ./src/

# Expose port
EXPOSE 8080

# Start server
CMD ["uvicorn", "m365_mcp.server:app", "--host", "0.0.0.0", "--port", "8080"]
