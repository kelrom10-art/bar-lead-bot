FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY index.html .
COPY manifest.json .
COPY icon-192.png .
COPY icon-512.png .

ENV PORT=8080
EXPOSE 8080

CMD ["python", "app.py"]
