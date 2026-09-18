FROM python:3.12-slim

# Cài đặt LibreOffice và Font chữ
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Cài đặt thư viện Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy toàn bộ code vào container
COPY . .

# Gom static files
RUN python manage.py collectstatic --no-input

EXPOSE 10000

# Chạy ứng dụng bằng Gunicorn (thay 'core.wsgi' bằng tên_project.wsgi của bạn)
CMD ["gunicorn", "--bind", "0.0.0.0:10000", "core.wsgi:application"]
