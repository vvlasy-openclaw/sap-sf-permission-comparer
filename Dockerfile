FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY server.py .
COPY comparer.py .
COPY frontend/ frontend/

ARG BUILD_DATE=unknown
ENV BUILD_DATE=${BUILD_DATE}

EXPOSE 8000

CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
