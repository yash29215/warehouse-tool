FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (layer-cached until requirements change)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

# Copy application code
COPY . .

# Runtime directories (overridden by volumes in docker-compose)
RUN mkdir -p uploads downloads

EXPOSE 5000

# Single worker + 8 threads — required to keep the in-memory SSE stream dict
# shared across all requests (multi-worker would break live progress bars)
CMD ["gunicorn", \
     "--bind",        "0.0.0.0:5000", \
     "--workers",     "1", \
     "--threads",     "8", \
     "--timeout",     "600", \
     "--access-logfile", "-", \
     "app:app"]
