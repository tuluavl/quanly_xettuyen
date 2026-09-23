# urls.py
from django.urls import path
from . import views

urlpatterns = [
    # Các đường dẫn DSTS
    path('', views.index, name='index'),
    path('import-thi-sinh/', views.import_thi_sinh, name='import_thi_sinh'),
    path('export-thi-sinh/', views.export_thi_sinh, name='export_thi_sinh'),
    
    # Các đường dẫn điểm chuẩn
    
    path('diem-chuan/', views.diem_chuan, name='diem_chuan'), # Thêm đường dẫn này
    path('export-diem-chuan/', views.export_diem_chuan, name='export_diem_chuan'),
    
    # Các đường dẫn tài khoản
    
    path('', views.index, name='index'),                     # CHỈ trang chủ mới dùng ''
    path('verify-otp/', views.verify_otp_view, name='verify_otp'),
    path('login/', views.login_view, name='login'),           # BẮT BUỘC phải có 'login/'
    path('logout/', views.logout_view, name='logout'),
    path('quan-ly-tai-khoan/', views.manage_users, name='manage_users'),

    
    
   # Các đường dẫn in giấy báo
    path('danh-sach-in-giay-bao/', views.ds_in_giay_bao, name='ds_in_giay_bao'),
    path('ds-in-giay-bao/', views.ds_in_giay_bao, name='ds_in_giay_bao'),
    path('in-giay-bao/<str:cccd>/', views.in_giay_bao, name='in_giay_bao'),
    path('tra-cuu/giay-bao-mat/', views.in_giay_bao_sinh_vien, name='in_giay_bao_sinh_vien'),
    path('quan-ly/luu-cau-hinh-giay-bao/', views.luu_cau_hinh_giay_bao, name='luu_cau_hinh_giay_bao'),
    path('export-all-pdf/start/', views.start_export_all_pdf, name='start_export_all_pdf'),
    path('export-all-pdf/status/<str:task_id>/', views.check_export_status, name='check_export_status'),
    path('export-all-pdf/download/<str:task_id>/', views.download_export_zip, name='download_export_zip'),
    
    # Các đường dẫn danh sach truong THPT
    path('truong-thpt/', views.danh_sach_truong_thpt, name='danh_sach_truong_thpt'),
    path('truong-thpt/xoa/<int:id>/', views.xoa_truong_thpt, name='xoa_truong_thpt'),
    path('truong-thpt/sua/<int:id>/', views.sua_truong_thpt, name='sua_truong_thpt'),
    path('truong-thpt/import/', views.import_to_hop_mon, name='import_truong_thpt'),
    path('truong-thpt/xuat-excel/', views.xuat_excel_truong_thpt, name='xuat_excel_truong_thpt'),    
    
    # Các đường dẫn THM
     path('to-hop-mon/', views.to_hop_mon, name='to_hop_mon'),
    path('to-hop-mon/import/', views.import_to_hop_mon, name='import_to_hop_mon'),
    path('to-hop-mon/export/', views.xuat_excel_to_hop_mon, name='xuat_excel_to_hop_mon'),
    path('to-hop-mon/sua/<str:ma_to_hop>/', views.sua_to_hop_mon, name='sua_to_hop_mon'),
    path('to-hop-mon/xoa/<str:ma_to_hop>/', views.xoa_to_hop_mon, name='xoa_to_hop_mon'),
    path('to-hop-mon/xoa-tat-ca/', views.xoa_tat_ca_to_hop_mon, name='xoa_tat_ca_to_hop_mon'),
    
    # Các đường dẫn VSAT
    path('diem-vsat/', views.danh_sach_diem_vsat, name='danh_sach_diem_vsat'),
    path('diem-vsat/import/', views.import_diem_vsat, name='import_diem_vsat'),
    path('diem-vsat/sua/<int:pk>/', views.sua_diem_vsat, name='sua_diem_vsat'),
    path('diem-vsat/xoa/<int:pk>/', views.xoa_diem_vsat, name='xoa_diem_vsat'),
    path('diem-vsat/export/', views.xuat_excel_diem_vsat, name='xuat_excel_diem_vsat'),
    
    # Các đường dẫn Danh sach trung tuyen
    path('danh-sach-trung-tuyen/', views.danh_sach_trung_tuyen, name='danh_sach_trung_tuyen'),
    path('tao-ma-sv/', views.tao_ma_sv, name='tao_ma_sv'),
    path('cap-nhat-so-cv/', views.cap_nhat_so_cv, name='cap_nhat_so_cv'),
    path('import-bo-sung/', views.import_bo_sung_trung_tuyen, name='import_bo_sung_trung_tuyen'),
    path('xuat-excel-trung-tuyen/', views.xuat_excel_trung_tuyen, name='xuat_excel_trung_tuyen'),
    path('export-sms/', views.export_sms, name='export_sms'),
    path('trung-tuyen/sua/<int:pk>/', views.sua_trung_tuyen, name='sua_trung_tuyen'),
    path('trung-tuyen/xoa/<int:pk>/', views.xoa_trung_tuyen, name='xoa_trung_tuyen'),
    
    # Các đường dẫn cập nhật bổ sung thông tin nhập học
    path('dong-bo-thong-tin-trung-tuyen/', views.dong_bo_thong_tin_trung_tuyen, name='dong_bo_thong_tin_trung_tuyen'),
    
    
    
   # Các đường dẫn export nhập học
    path('export/import-nhap-hoc/', views.export_excel_import_nhap_hoc, name='export_import_nhap_hoc'),
    path('export/hoso-trung-tuyen/', views.export_excel_hoso_trung_tuyen, name='export_hoso_trung_tuyen'),
    
   # Các đường dẫn cập nhật thông tin trúng tuyển  
    path('cap-nhat-trung-tuyen/', views.danh_sach_cap_nhat_trung_tuyen, name='danh_sach_cap_nhat_trung_tuyen'),
    path('cap-nhat-trung-tuyen/sua/<int:pk>/', views.sua_cap_nhat_trung_tuyen, name='sua_cap_nhat_trung_tuyen'),
    path('cap-nhat-trung-tuyen/xoa/<int:pk>/', views.xoa_cap_nhat_trung_tuyen, name='xoa_cap_nhat_trung_tuyen'),
    path('cap-nhat-trung-tuyen/export/', views.export_cap_nhat_trung_tuyen, name='export_cap_nhat_trung_tuyen'),
    path('cap-nhat-trung-tuyen/import/', views.danh_sach_cap_nhat_trung_tuyen, name='import_cap_nhat_trung_tuyen'),


    # Tuyến đường cho sinh viên tra cứu
    path('tra-cuu/', views.sinh_vien_login, name='sinh_vien_login'),
    path('tra-cuu/ket-qua/', views.sinh_vien_dashboard, name='sinh_vien_dashboard'),
    path('tra-cuu/logout/', views.sinh_vien_logout, name='sinh_vien_logout'),
    
    # Route in giấy báo bảo mật không tham số URL
    path('tra-cuu/xem-giay-bao-pdf/', views.in_giay_bao_bao_mat, name='in_giay_bao_bao_mat'),
    path('tra-cuu/xem-giay-bao-pdf/',views.in_giay_bao_sinh_vien, name='xem_giay_bao_sinh_vien',),


]