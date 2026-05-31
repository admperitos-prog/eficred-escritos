FROM python:3.11-slim

# Instalar LibreOffice para conversión a PDF
RUN apt-get update && apt-get install -y \
    libreoffice \
    libreoffice-writer \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 5000
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "app:app"]
