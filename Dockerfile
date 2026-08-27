FROM python:3.13-slim
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 E2_ENV=production E2_NO_BROWSER=1
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
EXPOSE 5000
CMD ["gunicorn","app:app","--bind","0.0.0.0:5000","--workers","2","--threads","4","--timeout","120"]
