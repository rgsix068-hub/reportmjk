# ============================================================
# Dockerfile — Daily Report Generator (Streamlit + Liberation Fonts)
# Deploy target: Railway.app
# ============================================================

FROM python:3.11-slim-bookworm

# ── System dependencies ─────────────────────────────────────
# - fonts-liberation2: metric-compatible with Arial/Times New Roman
# - libreoffice-writer: convert DOCX -> PDF with high fidelity
# - xvfb: virtual display server (dibutuhkan LibreOffice untuk render tabel/layout)
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        fonts-liberation2 \
        fontconfig \
        ca-certificates \
        libreoffice-writer \
        xvfb \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f -v

# ── Working directory ───────────────────────────────────────
WORKDIR /app

# LibreOffice membutuhkan HOME writable untuk config directory
ENV HOME=/app

# ── Python dependencies ─────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Application code ────────────────────────────────────────
COPY app.py .
COPY template.docx .
COPY template.pdf .

# ── Expose port ─────────────────────────────────────────────
EXPOSE 8080

# ── Run Streamlit ───────────────────────────────────────────
ENV STREAMLIT_SERVER_HEADLESS=true
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

ENTRYPOINT ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0"]
