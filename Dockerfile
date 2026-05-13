# syntax=docker/dockerfile:1.7
# ─────────────────────────────────────────────────────────────────────────────
# SpinCycle — multi-stage Dockerfile for the Streamlit-based QC pipeline.
#
# Build:   docker build -t spincycle:latest .
# Run:     docker run --rm -p 8501:8501 spincycle:latest
# Browse:  http://localhost:8501
#
# The image is a Linux/macOS alternative to the offline Windows-bundled
# deployment. The .bat workflow remains the supported path for air-gapped
# instrument PCs; this image exists for developers, cloud demos, and CI.
# ─────────────────────────────────────────────────────────────────────────────

# ── Stage 1: builder — install Python dependencies into a venv ───────────────
FROM python:3.11-slim-bookworm AS builder

# Don't write .pyc files, don't buffer stdout (so logs stream immediately).
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /build

# Copy ONLY the dependency manifests first so the pip-install layer caches
# unless requirements actually change. Source-code changes won't bust this.
COPY requirements.txt .

# Build a self-contained venv we can copy to the runtime stage.
RUN python -m venv /opt/venv \
    && /opt/venv/bin/pip install --upgrade pip \
    && /opt/venv/bin/pip install -r requirements.txt

# ── Stage 2: runtime — minimal image with just the venv + app code ──────────
FROM python:3.11-slim-bookworm AS runtime

# OCI image labels — surface project metadata in tooling like GHCR and Docker Desktop.
LABEL org.opencontainers.image.title="spincycle" \
      org.opencontainers.image.description="Offline-capable QC pipeline for Agilent Chemstation/MassHunter exports" \
      org.opencontainers.image.source="https://github.com/stefanhaugen/spincycle_sanitized" \
      org.opencontainers.image.licenses="MIT"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/opt/venv/bin:$PATH" \
    SPINCYCLE_LOG_LEVEL=INFO

# Create a non-root user. Running as root is a security anti-pattern even
# in dev — keeping the app un-privileged limits blast radius if exploited.
RUN groupadd --system spincycle \
    && useradd --system --gid spincycle --create-home --home-dir /home/spincycle spincycle

WORKDIR /app

# Copy the prebuilt venv from the builder stage (layer cache hit if deps
# unchanged) and the application source.
COPY --from=builder /opt/venv /opt/venv
COPY --chown=spincycle:spincycle app.py spincycle_utils.py spincycle_logging.py ./

# Drop privileges before runtime.
USER spincycle

# Streamlit's default port. Documented for tooling; doesn't actually
# publish the port — that's `docker run -p` or compose's job.
EXPOSE 8501

# Container healthcheck — Streamlit exposes a built-in /_stcore/health
# endpoint that returns 200 when the server is accepting requests.
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
    CMD python -c "import urllib.request,sys; \
sys.exit(0 if urllib.request.urlopen('http://localhost:8501/_stcore/health', timeout=3).status == 200 else 1)" \
    || exit 1

# Bind to 0.0.0.0 so the host can reach us via the published port. Disable
# Streamlit's anonymous usage stats since this runs in restricted/offline
# environments where outbound calls are unwelcome.
ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.address=0.0.0.0", \
            "--server.port=8501", \
            "--browser.gatherUsageStats=false"]
