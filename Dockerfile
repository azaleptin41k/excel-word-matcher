# ── Stage 1: Builder ─────────────────────────────────────────────────────────
FROM python:3.10-slim AS builder

# Prevent Python from writing .pyc and enable unbuffered output
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# Install only the system packages needed for the build
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       python3-tk \
       binutils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /build

# Install Python dependencies (cached layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy source code and build assets
COPY beaver2.py beaver2.spec categories.docx ./

# Build the binary
RUN pyinstaller --clean -y beaver2.spec

# ── Stage 2: Runtime (minimal) ──────────────────────────────────────────────
FROM debian:bookworm-slim AS runtime

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libpython3.11 \
       python3-tk \
    && rm -rf /var/lib/apt/lists/*

# Create non-root user
RUN groupadd --gid 1000 appuser \
    && useradd --uid 1000 --gid appuser --shell /bin/bash --create-home appuser

WORKDIR /app

# Copy built binary from builder stage
COPY --from=builder /build/dist/ ./dist/

RUN chown -R appuser:appuser /app

USER appuser

LABEL org.opencontainers.image.title="excel-word-matcher" \
      org.opencontainers.image.description="GUI utility for Excel-Word cross-matching" \
      org.opencontainers.image.source="https://github.com/azaleptin41k/excel-word-matcher"

CMD ["./dist/beaver2"]