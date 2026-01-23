FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Copy project files (README.md required for hatchling metadata)
COPY pyproject.toml README.md ./
COPY src/ src/

# Install Python dependencies
RUN pip install --no-cache-dir .

# Create non-root user for security
RUN useradd --create-home --shell /bin/bash appuser && \
    mkdir -p /home/appuser/.m365-mcp && \
    chown -R appuser:appuser /app /home/appuser

USER appuser

# Set environment
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1

# Run the MCP server
CMD ["python", "-m", "m365_mcp.server"]
