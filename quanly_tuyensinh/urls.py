from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('xettuyen.urls')),  # Trỏ toàn bộ đường dẫn về app xettuyen
]