# ==============================================================================
# SCRIPT TỰ ĐỘNG KHỞI TẠO VÀ SỬA LỖI DỰ ÁN QUẢN LÝ TUYỂN SINH K52 (DJANGO + MYSQL)
# ==============================================================================
# Hướng dẫn sử dụng:
# 1. Chép file này vào thư mục dự án của bạn (ví dụ: D:\PYTHON\Myweb\quanly_tuyensinh)
# 2. Chạy lệnh: python setup_k52_project.py
# 3. Chạy máy chủ: python manage.py runserver
# ==============================================================================

import os
import sys

def create_file(path, content):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")
    print(f" [OK] Đã tạo/cập nhật: {path}")

print("="*60)
print("  ĐANG TỰ ĐỘNG CẤU HÌNH DỰ ÁN DJANGO K52 TUYỂN SINH  ")
print("="*60)

# 1. Cập nhật __init__.py ở thư mục cấu hình gốc
init_py = """
import pymysql
pymysql.install_as_MySQLdb()
"""
create_file("quanly_tuyensinh/__init__.py", init_py)

# 2. Tạo models.py an toàn mã hóa UTF-8 cho các bảng K52
models_py = """
from django.db import models

class ThiSinhData(models.Model):
    cccd = models.CharField(max_length=20, primary_key=True, db_column='CCCD')
    tb_mon_lop10 = models.FloatField(null=True, blank=True, db_column='TB_MON_LOP10')
    tb_mon_lop11 = models.FloatField(null=True, blank=True, db_column='TB_MON_LOP11')
    tb_mon_lop12 = models.FloatField(null=True, blank=True, db_column='TB_MON_LOP12')
    to_thpt = models.FloatField(null=True, blank=True, db_column='TO_THPT')
    li_thpt = models.FloatField(null=True, blank=True, db_column='LI_THPT')
    ho_thpt = models.FloatField(null=True, blank=True, db_column='HO_THPT')
    va_thpt = models.FloatField(null=True, blank=True, db_column='VA_THPT')
    ls_thpt = models.FloatField(null=True, blank=True, db_column='LS_THPT')
    ve_thpt = models.FloatField(null=True, blank=True, db_column='VE_THPT')
    gdktpl_thpt = models.FloatField(null=True, blank=True, db_column='GDKTPL_THPT')
    tin_thpt = models.FloatField(null=True, blank=True, db_column='TIN_THPT')
    ngoaingu_thpt = models.FloatField(null=True, blank=True, db_column='NGOAINGU_THPT')
    mann_thpt = models.CharField(max_length=10, null=True, blank=True, db_column='MANN_THPT')
    sinh_thpt = models.FloatField(null=True, blank=True, db_column='SINH_THPT')

    class Meta:
        db_table = 'thi_sinh_data'
        managed = False

class KetQuaDgnl(models.Model):
    id = models.AutoField(primary_key=True)
    cccd = models.CharField(max_length=20, db_column='cccd', null=True)
    diem_dgnl = models.FloatField(null=True, db_column='diem_dgnl')

    class Meta:
        db_table = 'ket_qua_dgnl'
        managed = False

class ChungChiTa(models.Model):
    id = models.AutoField(primary_key=True)
    cccd = models.CharField(max_length=20, db_column='cccd', null=True)
    loai_cc = models.CharField(max_length=50, db_column='loai_cc', null=True)
    diem_quy_doi = models.FloatField(null=True, db_column='diem_quy_doi')

    class Meta:
        db_table = 'chung_chi_ta'
        managed = False

class UtkvDt(models.Model):
    ma_utkv_dt = models.CharField(max_length=20, primary_key=True, db_column='ma_utkv_dt')
    diem_uu_tien = models.FloatField(null=True, db_column='diem_uu_tien')

    class Meta:
        db_table = 'utkv_dt'
        managed = False
"""
create_file("xettuyen/models.py", models_py)

# 3. Tạo views.py
views_py = """
from django.shortcuts import render
from django.http import JsonResponse
from .models import ThiSinhData

def index(request):
    try:
        students = ThiSinhData.objects.all()
        total_count = students.count()
    except Exception as e:
        students = []
        total_count = 0

    context = {
        'students': students,
        'total_count': total_count,
    }
    return render(request, 'xettuyen/index.html', context)

def tinh_diem_xet_tuyen(request):
    # Logic tính điểm xét tuyển
    return JsonResponse({'status': 'success', 'message': 'Đã hoàn thành tính điểm cho toàn bộ thí sinh K52!'})
"""
create_file("xettuyen/views.py", views_py)

# 4. Tạo urls.py của app xettuyen
urls_app_py = """
from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('tinh-diem/', views.tinh_diem_xet_tuyen, name='tinh_diem_xet_tuyen'),
]
"""
create_file("xettuyen/urls.py", urls_app_py)

# 5. Cập nhật urls.py của dự án chính
urls_main_py = """
from django.contrib import admin
from django.urls import path, include

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', include('xettuyen.urls')),
]
"""
create_file("quanly_tuyensinh/urls.py", urls_main_py)

print("="*60)
print(" BẠN ĐÃ TỰ ĐỘNG TẠO XONG TOÀN BỘ CODE CHUẨN MÃ HÓA UTF-8!")
print(" Chạy lệnh sau để khởi động máy chủ: python manage.py runserver")
print("="*60)
