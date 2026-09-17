import os
from pathlib import Path
import pymysql
from dotenv import load_dotenv  # Tải thư viện đọc biến môi trường

# Nạp dữ liệu từ file .env
load_dotenv()

pymysql.install_as_MySQLdb()

BASE_DIR = Path(__file__).resolve().parent.parent

# Bảo mật Secret Key qua biến môi trường (fallback giữ key mặc định nếu chưa cấu hình .env)
SECRET_KEY = os.getenv('SECRET_KEY', 'django-insecure-k52-tuyensinh-secret-key')
DEBUG = True
ALLOWED_HOSTS = ['*']

# Cấu hình chuyển hướng đăng nhập chuẩn
LOGIN_URL = 'login'
LOGIN_REDIRECT_URL = 'index'
LOGOUT_REDIRECT_URL = 'login'

INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'xettuyen',
]

STATIC_URL = '/static/'

STATICFILES_DIRS = [
    BASE_DIR / 'xettuyen' / 'static',
]

STATIC_ROOT = BASE_DIR / 'staticfiles'

MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'whitenoise.middleware.WhiteNoiseMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
    'xettuyen.middleware.CheckUserStatusMiddleware',
]

ROOT_URLCONF = 'quanly_tuyensinh.urls'

TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

WSGI_APPLICATION = 'quanly_tuyensinh.wsgi.application'

DATABASES = {
    'default': {
        'ENGINE': os.getenv('DB_ENGINE', 'django.db.backends.mysql'),
        'NAME': os.getenv('DB_NAME', 'k52'),
        'USER': os.getenv('DB_USER', 'root'),
        'PASSWORD': os.getenv('DB_PASSWORD', ''),
        'HOST': os.getenv('DB_HOST', 'localhost'),
        'PORT': os.getenv('DB_PORT', '3306'),
        'OPTIONS': {
            'charset': 'utf8mb4',
        },
    }
}

CACHES = {
    'default': {
        'BACKEND': 'django.core.cache.backends.locmem.LocMemCache',
        'LOCATION': 'unique-snowflake',
    }
}

# Cấu hình gửi mail qua Gmail SMTP (Đọc từ file .env)
EMAIL_BACKEND = 'django.core.mail.backends.smtp.EmailBackend'
EMAIL_HOST = 'smtp.gmail.com'
EMAIL_PORT = 587
EMAIL_USE_TLS = True
EMAIL_HOST_USER = os.getenv('EMAIL_HOST_USER', 'ntbinh82@gmail.com')
EMAIL_HOST_PASSWORD = os.getenv('EMAIL_HOST_PASSWORD', 'wticxjvlfdhfrxvz')
DEFAULT_FROM_EMAIL = f'Hệ Thống Xét Tuyển Bổ Sung UEH Mekong <{EMAIL_HOST_USER}>'

# Thời gian hết hạn Session tính bằng GIÂY (30 phút)
SESSION_COOKIE_AGE = 1800  

# Reset lại đồng hồ đếm ngược mỗi khi user thao tác
SESSION_SAVE_EVERY_REQUEST = True

# Đăng xuất ngay khi đóng trình duyệt web
SESSION_EXPIRE_AT_BROWSER_CLOSE = True

LANGUAGE_CODE = 'vi-vn'
TIME_ZONE = 'Asia/Ho_Chi_Minh'
USE_I18N = True
USE_TZ = True

DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'