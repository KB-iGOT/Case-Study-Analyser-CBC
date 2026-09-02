FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

# Install required system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Install uv
RUN pip install --no-cache-dir uv

# Copy dependency files first for better Docker layer caching
COPY pyproject.toml uv.lock ./

# Install Python dependencies
RUN uv sync --frozen

# Copy application files
COPY app.py .
COPY assessment_criteria.py .
COPY case_generator.py .
COPY db_models.py .
COPY utils.py .

# Copy required application directories
COPY .streamlit/ .streamlit/
COPY attached_assets/ attached_assets/
COPY docs/ docs/
COPY scripts/ scripts/

EXPOSE 5000

CMD ["uv", "run", "streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=5000"]