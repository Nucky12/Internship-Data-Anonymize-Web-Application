FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 PIP_NO_CACHE_DIR=1 HOST=0.0.0.0 PORT=5000

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY anonymizer.py app.py taxonomy.json ./
COPY templates templates

RUN useradd -m app
USER app

EXPOSE 5000
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:5000/health')"

CMD ["python", "app.py"]
