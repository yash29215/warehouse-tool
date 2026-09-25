FROM python:3.11-slim

WORKDIR /app

# Install dependencies first (layer-cached until requirements change)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

# ── ODA File Converter (DWG → DXF, for the "DWG → Map" tab) ──────────────
# No macOS build exists, which is why this only happens inside this Linux
# image rather than directly on a Mac host. Off by default — normal builds
# (`docker build .`) are completely unaffected. To enable:
#   1. Download the Linux .deb yourself from
#      https://www.opendesign.com/guestfiles/oda_file_converter
#      (requires accepting ODA's license on their site — can't be scripted)
#   2. Save it as vendor/ODAFileConverter.deb
#   3. docker build --build-arg INSTALL_ODA=true .
ARG INSTALL_ODA=false
COPY vendor/ /tmp/vendor/
RUN if [ "$INSTALL_ODA" = "true" ] && [ -f /tmp/vendor/ODAFileConverter.deb ]; then \
      apt-get update && \
      apt-get install -y --no-install-recommends \
        xvfb xauth libgl1 libxkbcommon0 libglib2.0-0 libfontconfig1 \
        libxkbcommon-x11-0 libxcb-cursor0 libxcb-icccm4 libxcb-image0 \
        libxcb-keysyms1 libxcb-randr0 libxcb-render-util0 libxcb-shape0 \
        libxcb-sync1 libxcb-xfixes0 libxcb-xinerama0 libxcb-xkb1 && \
      apt-get install -y --no-install-recommends /tmp/vendor/ODAFileConverter.deb && \
      rm -rf /var/lib/apt/lists/*; \
    fi && rm -rf /tmp/vendor

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
