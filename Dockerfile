FROM python:3.12-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-core fonts-dejavu fonts-noto-core fonts-noto-cjk && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 8080
CMD ["gunicorn","-w","2","-b","0.0.0.0:8080","api.index:app","--timeout","120"]
