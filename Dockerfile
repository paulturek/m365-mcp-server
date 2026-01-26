FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy project files
COPY pyproject.toml .
COPY README.md .
COPY src/ src/

# Install Python dependencies
RUN pip install --no-cache-dir -e .

# Create non-root user for security
RUN useradd --create-home --shell /bin/bash appuser && \
    mkdir -p /home/appuser/.m365-mcp && \
    chown -R appuser:appuser /app /home/appuser

USER appuser

# Expose port for Railway
EXPOSE 8000

# Run the MCP server (auto-detects Railway vs local)
CMD ["python", "-m", "m365_mcp"]
