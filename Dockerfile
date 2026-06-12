# ============================================================
# Dockerfile — Daily Report Generator (Streamlit + Arial fonts)
# Deploy target: Railway.app
# ============================================================

FROM python:3.11-slim

# ── System dependencies ─────────────────────────────────────
# Install Arial + Times New Roman fonts with EULA auto-accept
RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections \
    && apt-get update \
    && apt-get install -y --no-install-recommends \
        ttf-mscorefonts-installer \
        fontconfig \
        ca-certificates \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -f -v

# ── Working directory ───────────────────────────────────────
WORKDIR /app

# ── Python dependencies ─────────────────────────────────────
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# ── Application code ────────────────────────────────────────
COPY app.py .
COPY template.docx .
COPY template.pdf .

# ── Expose port ─────────────────────────────────────────────
EXPOSE 8080

# ── Health check ────────────────────────────────────────────
HEALTHCHECK --interval=30s --timeout=10s --start-period=30s --retries=3 \
    CMD python3 -c "import urllib.request; urllib.request.urlopen('http://localhost:${PORT:-8080}/_stcore/health')" || exit 1

# ── Run Streamlit ───────────────────────────────────────────
ENV STREAMLIT_SERVER_HEADLESS=true
ENV STREAMLIT_SERVER_ENABLE_CORS=false
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

ENTRYPOINT ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8080} --server.address=0.0.0.0"]
