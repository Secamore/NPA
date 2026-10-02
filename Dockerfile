FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    NPA_DB_PATH=/data/npa_portal.db \
    NPA_ASSETS_DIR=/data/assets

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .
RUN useradd -m npa && mkdir -p /data/assets && chown -R npa /app /data
USER npa

EXPOSE 8501
VOLUME /data

HEALTHCHECK --interval=30s --timeout=5s CMD python -c "import os,urllib.request as u; u.urlopen('http://localhost:%s/_stcore/health' % os.environ.get('PORT','8501'))"

# Hosts such as Render/Railway/Fly inject $PORT; locally it falls back to 8501.
CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8501} --server.address=0.0.0.0"]
