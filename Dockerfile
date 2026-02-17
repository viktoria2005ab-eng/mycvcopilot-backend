FROM python:3.11-slim

# 1) Dépendances système + LibreOffice + polices (important pour le rendu)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-core \
    fonts-dejavu \
    fonts-liberation \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# 2) Dossier de travail
WORKDIR /app

# 3) Python deps
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# 4) Code
COPY . /app

# 5) Render fournit le PORT en env
ENV PORT=10000

# 6) Lancer FastAPI
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT}"]
