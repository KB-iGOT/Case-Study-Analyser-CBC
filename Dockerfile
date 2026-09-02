# syntax=docker/dockerfile:1

FROM python:3.11-slim

# Prevent Python from writing .pyc files
# and ensure logs are sent directly to stdout/stderr
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Install only required OS packages
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy dependency files first for Docker layer caching
COPY pyproject.toml uv.lock ./

# Install uv
RUN pip install --no-cache-dir uv

# Install Python dependencies from uv.lock
RUN uv sync --frozen

# Copy application source
COPY app.py .
COPY assessment_criteria.py .
COPY case_generator.py .
COPY db_models.py .
COPY utils.py .

# Copy application resources
COPY .streamlit/ .streamlit/
COPY attached_assets/ attached_assets/
COPY docs/ docs/
COPY scripts/ scripts/

# Streamlit port
EXPOSE 5000

# Streamlit configuration
ENV STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
    STREAMLIT_SERVER_PORT=5000 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

CMD ["uv", "run", "streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=5000"]