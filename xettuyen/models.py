from django.db import models
from django.utils import timezone
from django.conf import settings

class Role(models.Model):
    code = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=255)
    permissions = models.ManyToManyField('Permission', through='RolePermission', related_name='roles')

    class Meta:
        db_table = 'roles'

    def __str__(self):
        return self.name


class Permission(models.Model):
    GROUP_CHOICES = [
        ('thisinh', 'Danh sách thí sinh'),
        ('to_hop_mon', 'Tổ hợp môn'),
        ('truong_thpt', 'Danh sách trường THPT'),
        ('diem_vsat', 'Danh sách điểm VSAT'),
        ('diem_chuan', 'Điểm chuẩn'),
        ('cap_nhat_trung_tuyen', 'Cập nhật thông tin trúng tuyển'),
        ('ds_trung_tuyen', 'Danh sách trúng tuyển'),
        ('in_giay_bao', 'In giấy báo'),
        ('system', 'Quản lý hệ thống'),
        ('other', 'Chức năng khác'),
    ]

    name = models.CharField(max_length=255)
    code = models.CharField(max_length=100, unique=True)
    group = models.CharField(max_length=50, choices=GROUP_CHOICES, default='other')

    class Meta:
        db_table = 'permissions'

    def __str__(self):
        return f"{self.name} ({self.code})"


class RolePermission(models.Model):
    role = models.ForeignKey(Role, on_delete=models.CASCADE, db_column='role_id')
    permission = models.ForeignKey(Permission, on_delete=models.CASCADE, db_column='permission_id')

    class Meta:
        db_table = 'role_permissions'
        unique_together = ('role', 'permission')


class CustomUser(models.Model):
    username = models.CharField(max_length=50, unique=True)
    password = models.CharField(max_length=255)
    full_name = models.CharField(max_length=100)
    role = models.ForeignKey(Role, on_delete=models.SET_NULL, null=True, blank=True, db_column='role_id')
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    email = models.EmailField(max_length=255, null=True, blank=True)
    class Meta:
        db_table = 'users'

    def __str__(self):
        return self.full_name or self.username

    def has_menu_perm(self, perm_code):
        if not self.role:
            return False
        if self.role.code == 'admin' or self.role.name == 'Quản trị viên':
            return True
        return self.role.permissions.filter(code=perm_code).exists()


class UserProfile(models.Model):
    # Thay User thành CustomUser
    user = models.OneToOneField(CustomUser, on_delete=models.CASCADE, related_name='profile')
    session_version = models.IntegerField(default=1)  # Tăng số này để Force Logout
    otp_secret = models.CharField(max_length=32, blank=True, null=True)
    is_2fa_enabled = models.BooleanField(default=False)


class AuditLog(models.Model):
    action = models.CharField(max_length=50)
    details = models.TextField(null=True, blank=True)
    ip_address = models.CharField(max_length=39, null=True, blank=True)
    created_at = models.DateTimeField(auto_now_add=True)

    # Thay settings.AUTH_USER_MODEL bằng 'CustomUser'
    actor = models.ForeignKey(
        'CustomUser',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='actor_audit_logs'
    )
    target_user = models.ForeignKey(
        'CustomUser',
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name='target_audit_logs'
    )

    class Meta:
        ordering = ['-created_at']


class KetQuaLocAo(models.Model):
    ma_xet_tuyen_moi = models.CharField(max_length=50, primary_key=True, db_column='ma_xet_tuyen_moi')
    ten_ma_xet_tuyen = models.TextField(blank=True, null=True, db_column='ten_ma_xet_tuyen')
    ma_nganh = models.TextField(blank=True, null=True, db_column='ma_nganh')
    chi_tieu_chung = models.TextField(blank=True, null=True, db_column='chi_tieu_chung')
    diem_chuan = models.TextField(blank=True, null=True, db_column='diem_chuan')
    sl_tt = models.TextField(blank=True, null=True, db_column='sl_tt')
    ti_le_tt_ct = models.TextField(blank=True, null=True, db_column='ti_le_tt_ct')
    nv1 = models.TextField(blank=True, null=True, db_column='NV1')
    nv2 = models.TextField(blank=True, null=True, db_column='NV2')
    nv3 = models.TextField(blank=True, null=True, db_column='NV3')

    class Meta:
        db_table = 'ket_qua_loc_ao'
        managed = True


class DanhMucNganh(models.Model):
    ma_nganh = models.CharField(max_length=20, unique=True, verbose_name="Mã ngành")
    ten_nganh = models.CharField(max_length=255, verbose_name="Tên ngành")
    chi_tieu = models.IntegerField(default=0, verbose_name="Chỉ tiêu")
    diem_dau = models.FloatField(default=60.0, verbose_name="Điểm đậu")

    def __str__(self):
        return f"{self.ma_nganh} - {self.ten_nganh}"


class ThiSinhData(models.Model):
    stt = models.TextField(primary_key=True) 
    sbd = models.TextField(blank=True, null=True)
    ho_ten = models.TextField(blank=True, null=True)
    cccd = models.TextField(db_column='CCCD', blank=True, null=True)
    dtut = models.TextField(blank=True, null=True)
    kvut = models.TextField(blank=True, null=True)
    nam_tn_thpt = models.TextField(blank=True, null=True)
    hoc_luc_ket_qua_hoc_tap = models.TextField(blank=True, null=True)
    diem_tb_cac_nam_hoc = models.TextField(blank=True, null=True)
    noi_thuong_tru_ma_tinh = models.TextField(blank=True, null=True)
    ma_tinh_lop_12 = models.TextField(blank=True, null=True)
    ma_truong_lop_12 = models.TextField(blank=True, null=True)
    to = models.TextField(blank=True, null=True)
    va = models.TextField(blank=True, null=True)
    li = models.TextField(blank=True, null=True)
    ho = models.TextField(blank=True, null=True)
    si = models.TextField(blank=True, null=True)
    su = models.TextField(blank=True, null=True)
    di = models.TextField(blank=True, null=True)
    gdcd = models.TextField(blank=True, null=True)
    nn = models.TextField(blank=True, null=True)
    ma_mon_nn = models.TextField(blank=True, null=True)
    ktpl = models.TextField(blank=True, null=True)
    ti = models.TextField(blank=True, null=True)
    cncn = models.TextField(blank=True, null=True)
    cnnn = models.TextField(blank=True, null=True)
    nk1 = models.TextField(blank=True, null=True)
    diem_xet_tot_nghiep = models.TextField(blank=True, null=True)
    dan_toc = models.TextField(blank=True, null=True)
    ma_dan_toc = models.TextField(blank=True, null=True)
    ma_tinh_lop_10 = models.TextField(blank=True, null=True)
    ma_truong_lop_10 = models.TextField(blank=True, null=True)
    ma_tinh_lop_11 = models.TextField(blank=True, null=True)
    ma_truong_lop_11 = models.TextField(blank=True, null=True)
    diem_dtut_30 = models.TextField(blank=True, null=True)
    diem_kvut_30 = models.TextField(blank=True, null=True)
    malop12 = models.TextField(blank=True, null=True)
    malop10 = models.TextField(blank=True, null=True)
    malop11 = models.TextField(blank=True, null=True)
    cong_chuyen = models.TextField(blank=True, null=True)
    mon_dat_giai = models.TextField(blank=True, null=True)
    cap = models.TextField(blank=True, null=True)
    hang = models.TextField(blank=True, null=True)
    cong_giai_hsg = models.TextField(blank=True, null=True)
    diem_xet_thuong = models.TextField(blank=True, null=True)
    ccta = models.TextField(blank=True, null=True)
    cong_ccta = models.TextField(blank=True, null=True)
    a00_to_li_ho = models.TextField(blank=True, null=True)
    a01_to_li_n1 = models.TextField(blank=True, null=True)
    d01_to_va_n1 = models.TextField(blank=True, null=True)
    d07_to_ho_n1 = models.TextField(blank=True, null=True)
    d09_to_su_n1 = models.TextField(blank=True, null=True)
    d14_va_su_n1 = models.TextField(blank=True, null=True)
    x25_to_ktpt_n1 = models.TextField(blank=True, null=True)
    x26_to_ti_n1 = models.TextField(blank=True, null=True)
    v00_to_nk1_n1 = models.TextField(blank=True, null=True)
    to_va_li = models.TextField(blank=True, null=True)
    to_va_ho = models.TextField(blank=True, null=True)
    to_va_si = models.TextField(blank=True, null=True)
    to_va_su = models.TextField(blank=True, null=True)
    to_va_di = models.TextField(blank=True, null=True)
    to_va_gdcd = models.TextField(blank=True, null=True)
    to_va_ktpl = models.TextField(blank=True, null=True)
    to_va_cncn = models.TextField(blank=True, null=True)
    to_va_cnnn = models.TextField(blank=True, null=True)
    to_va_ti = models.TextField(blank=True, null=True)
    nv1 = models.TextField(blank=True, null=True)
    nv2 = models.TextField(blank=True, null=True)
    nv3 = models.TextField(blank=True, null=True)
    nguon_xt_thm_va_bat_ky = models.TextField(blank=True, null=True)
    nguon_xt_theo_thm_ko_co_v00 = models.TextField(blank=True, null=True)
    ket_hop_nguon_xt = models.TextField(blank=True, null=True)
    thi_sinh_xet_luat = models.TextField(blank=True, null=True)
    dieu_kien_nganh_luat = models.TextField(blank=True, null=True)
    nv1_1 = models.TextField(blank=True, null=True)
    to_hop_nv1 = models.TextField(blank=True, null=True)
    tong_diem_qd_nv1 = models.TextField(blank=True, null=True)
    nv2_1 = models.TextField(blank=True, null=True)
    to_hop_nv2 = models.TextField(blank=True, null=True)
    tong_diem_qd_nv2 = models.TextField(blank=True, null=True)
    nv3_1 = models.TextField(blank=True, null=True)
    to_hop_nv3 = models.TextField(blank=True, null=True)
    tong_diem_qd_nv3 = models.TextField(blank=True, null=True)
    diem_to_hop = models.TextField(blank=True, null=True)
    to_hop = models.TextField(blank=True, null=True)
    dotthi = models.TextField(blank=True, null=True)
    diem_DGNL = models.TextField(blank=True, null=True)
    DGNL_thang_30_quy_doi = models.TextField(blank=True, null=True)
    DGNL_100_nhan_06 = models.TextField(blank=True, null=True)
    thm_vsat = models.TextField(blank=True, null=True)
    diem_VSAT = models.TextField(blank=True, null=True)
    VSAT_thang_30_quy_doi = models.TextField(blank=True, null=True)
    VSAT_100_nhan_06 = models.TextField(blank=True, null=True)
    diem_max = models.TextField(blank=True, null=True)
    ky_thi = models.TextField(blank=True, null=True)
    diem_luat = models.TextField(blank=True, null=True)
    ky_thi_luat = models.TextField(blank=True, null=True)   
    diem_thi_max_qd100_nhan_06 = models.TextField(blank=True, null=True)
    diem_hoc_ba_qd100_nhan_04 = models.TextField(blank=True, null=True)
    diem_thi_hoc_ba = models.TextField(blank=True, null=True)
    diem_thi_hoc_ba_diem_cong = models.TextField(blank=True, null=True)
    diem_utdt = models.TextField(blank=True, null=True)
    diem_utkv = models.TextField(blank=True, null=True)
    diem_utkv_chua_giam_tru = models.TextField(blank=True, null=True)
    diem_utkv_giam_tru = models.TextField(blank=True, null=True)
    diem_xet_tuyen = models.TextField(blank=True, null=True)
    kq_NV1 = models.TextField(blank=True, null=True)
    ma_NV1 = models.TextField(blank=True, null=True)
    nv_NV1 = models.TextField(blank=True, null=True)
    kq_NV2 = models.TextField(blank=True, null=True)
    ma_NV2 = models.TextField(blank=True, null=True)
    nv_NV2 = models.TextField(blank=True, null=True)
    kq_NV3 = models.TextField(blank=True, null=True)
    ma_NV3 = models.TextField(blank=True, null=True)
    nv_NV3 = models.TextField(blank=True, null=True)
    kq_TT = models.TextField(blank=True, null=True)
    ma_TT = models.TextField(blank=True, null=True)
    nv_TT = models.TextField(blank=True, null=True)
    diem_thi_TT = models.TextField(blank=True, null=True)

    class Meta:
        managed = True
        db_table = 'thi_sinh_data'


class TruongTHPT(models.Model):
    ma_truong_ghep = models.CharField(max_length=20, blank=True, null=True, verbose_name="Mã trường ghép")
    ma_tinh = models.CharField(max_length=10, blank=True, null=True, verbose_name="Mã tỉnh")
    ten_tinh = models.CharField(max_length=150, blank=True, null=True, verbose_name="Tên tỉnh/TP")
    ma_phuong_xa = models.CharField(max_length=20, blank=True, null=True, verbose_name="Mã phường/xã")
    ten_xa_phuong = models.CharField(max_length=150, blank=True, null=True, verbose_name="Tên phường/xã")
    ma_truong = models.CharField(max_length=20, blank=True, null=True, verbose_name="Mã trường")
    ten_truong = models.CharField(max_length=255, blank=True, null=True, verbose_name="Tên trường")
    dia_chi = models.TextField(blank=True, null=True, verbose_name="Địa chỉ")
    khu_vuc = models.CharField(max_length=20, blank=True, null=True, verbose_name="Khu vực")

    class Meta:
        db_table = 'truong_thpt'
        verbose_name = 'Trường THPT'
        verbose_name_plural = 'Danh sách Trường THPT'


class DiemThiVsat(models.Model):
    so_cccd = models.CharField(max_length=20, verbose_name="Số CCCD")
    ho_ten = models.CharField(max_length=255, verbose_name="Họ và tên")
    
    di_vs = models.FloatField(null=True, blank=True)
    ho_vs = models.FloatField(null=True, blank=True)
    li_vs = models.FloatField(null=True, blank=True)
    n1_vs = models.FloatField(null=True, blank=True)
    si_vs = models.FloatField(null=True, blank=True)
    su_vs = models.FloatField(null=True, blank=True)
    to_vs = models.FloatField(null=True, blank=True)
    va_vs = models.FloatField(null=True, blank=True)
    max_score = models.FloatField(null=True, blank=True)
    
    thmon_a00_vsat = models.FloatField(null=True, blank=True)
    quy_doi_a00 = models.FloatField(null=True, blank=True)
    thmon_a01_vsat = models.FloatField(null=True, blank=True)
    quy_doi_a01 = models.FloatField(null=True, blank=True)
    thmon_d01_vsat = models.FloatField(null=True, blank=True)
    quy_doi_d01 = models.FloatField(null=True, blank=True)
    thmon_d07_vsat = models.FloatField(null=True, blank=True)
    quy_doi_d07 = models.FloatField(null=True, blank=True)
    thmon_d09_vsat = models.FloatField(null=True, blank=True)
    quy_doi_d09 = models.FloatField(null=True, blank=True)
    thmon_d14_vsat = models.FloatField(null=True, blank=True)
    quy_doi_d14 = models.FloatField(null=True, blank=True)
    diem_thi_vsat_max = models.FloatField(null=True, blank=True)

    class Meta:
        db_table = 'diem_thi_vsat'
        verbose_name = 'Điểm Thi V-SAT'
        verbose_name_plural = 'Danh sách Điểm Thi V-SAT'


class MauImportGiayBao(models.Model):
    IDSV = models.CharField(max_length=50, null=True, blank=True)
    so_cv = models.CharField(max_length=50, null=True, blank=True)
    ho_ten = models.CharField(max_length=150, null=True, blank=True)
    email = models.CharField(max_length=150, null=True, blank=True)
    dien_thoai = models.CharField(max_length=20, null=True, blank=True, verbose_name="Điện thoại")
    ma_dkxt = models.CharField(max_length=50, null=True, blank=True)
    ngay_sinh = models.CharField(max_length=20, null=True, blank=True)
    dtut = models.CharField(max_length=20, null=True, blank=True)
    kvut = models.CharField(max_length=20, null=True, blank=True)
    hoc_ba = models.CharField(max_length=255, null=True, blank=True)
    phuong_thuc_xet = models.CharField(max_length=50, null=True, blank=True)
    ptxt = models.CharField(max_length=100, null=True, blank=True)
    ctdt = models.CharField(max_length=255, null=True, blank=True)
    pt2_diem_cong = models.CharField(max_length=100, null=True, blank=True)
    pt2_tn_thm = models.CharField(max_length=100, null=True, blank=True)
    pt2_tn_mamon1 = models.CharField(max_length=20, null=True, blank=True)
    pt2_tn_diemmon1 = models.DecimalField(max_digits=4, decimal_places=2, null=True, blank=True)
    pt2_tn_mamon2 = models.CharField(max_length=20, null=True, blank=True)
    pt2_tn_diemmon2 = models.DecimalField(max_digits=4, decimal_places=2, null=True, blank=True)
    pt2_tn_mamon3 = models.CharField(max_length=20, null=True, blank=True)
    pt2_tn_diemmon3 = models.DecimalField(max_digits=4, decimal_places=2, null=True, blank=True)
    pt2_dgnl = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    pt2_vsat_thm = models.CharField(max_length=100, null=True, blank=True)
    pt2_vsat_mamon1 = models.CharField(max_length=20, null=True, blank=True)
    pt2_vsat_diemmon1 = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    pt2_vsat_mamon2 = models.CharField(max_length=20, null=True, blank=True)
    pt2_vsat_diemmon2 = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    pt2_vsat_mamon3 = models.CharField(max_length=20, null=True, blank=True)
    pt2_vsat_diemmon3 = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    pt2a_diemtbthpt = models.DecimalField(max_digits=5, decimal_places=2, null=True, blank=True)
    pt2_diem_qd = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    pt2_qd = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    cccd = models.CharField(max_length=20, null=True, blank=True)
    ma_sv = models.CharField(max_length=50, null=True, blank=True)
    barcode = models.CharField(max_length=50, null=True, blank=True)
    dtc0 = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    dc = models.DecimalField(max_digits=6, decimal_places=2, null=True, blank=True)
    page = models.IntegerField(null=True, blank=True)

    class Meta:
        db_table = 'mau_import_giay_bao'
        managed = True

    def __str__(self):
        return f"{self.ho_ten} - {self.cccd}"
        
class ToHopMon(models.Model): # Thay tên class theo đúng file models.py của bạn
    stt = models.IntegerField(default=0, null=True, blank=True)
    ma_to_hop_mon = models.CharField(max_length=10)
    ten_to_hop_mon = models.CharField(max_length=255)
    ma_mon_thi = models.CharField(max_length=255)

    class Meta:
        db_table = 'thm'  # Cố định tên bảng dưới database là 'thm'

class CapNhatThongTinTrungTuyen(models.Model):
    sbd = models.CharField(max_length=50, null=True, blank=True, verbose_name="Số báo danh")
    ma_noi_sinh = models.CharField(max_length=50, null=True, blank=True, verbose_name="Mã nơi sinh")
    ten_giai_hsg = models.CharField(max_length=255, null=True, blank=True, verbose_name="Tên giải HSG")
    ten_mon_hsg = models.CharField(max_length=100, null=True, blank=True, verbose_name="Tên môn HSG")
    ten_hang_hsg = models.CharField(max_length=100, null=True, blank=True, verbose_name="Tên hạng HSG")
    nam_hsg = models.CharField(max_length=10, null=True, blank=True, verbose_name="Năm đạt giải HSG")
    cccd = models.CharField(max_length=20, db_index=True, verbose_name="Số CCCD")
    ma_dkxt = models.CharField(max_length=50, blank=True, null=True, verbose_name="Mã ĐKXT")
    ngay_sinh = models.DateField(blank=True, null=True, verbose_name="Ngày sinh")
    gioi_tinh = models.CharField(max_length=10, blank=True, null=True, verbose_name="Giới tính")
    dan_toc = models.CharField(max_length=50, blank=True, null=True, verbose_name="Dân tộc")
    email_sv = models.EmailField(blank=True, null=True, verbose_name="Email SV")
    dien_thoai= models.CharField(max_length=20, null=True, blank=True, verbose_name="Điện thoại")
    mssv = models.CharField(max_length=20, blank=True, null=True, verbose_name="MSSV")
    mat_khau_online = models.CharField(max_length=100, blank=True, null=True, verbose_name="Mật khẩu Online")
    email_ueh = models.EmailField(blank=True, null=True, verbose_name="Email UEH")
    mat_khau_email = models.CharField(max_length=100, blank=True, null=True, verbose_name="Mật khẩu Email")
    tong_diem = models.FloatField(blank=True, null=True, verbose_name="Tổng điểm")
    dtc0_pt1 = models.FloatField(blank=True, null=True)
    dtc0_pt2 = models.FloatField(blank=True, null=True)
    loai_ccqt_uid = models.CharField(max_length=100, blank=True, null=True, verbose_name="Loại chứng chỉ QT")
    so_diem_ccqt = models.FloatField(blank=True, null=True)
    ngay_thi_ccqt = models.DateField(blank=True, null=True)
    ngay_thi_ccqt_new = models.DateField(blank=True, null=True)
    diem_tb_lop10 = models.FloatField(blank=True, null=True)
    diem_tb_lop11 = models.FloatField(blank=True, null=True)
    diem_tb_lop12 = models.FloatField(blank=True, null=True)

    class Meta:
        db_table = 'cap_nhat_thong_tin_trung_tuyen'
        verbose_name = 'Cập nhật thông tin trúng tuyển'


class CauHinhGiayBao(models.Model):
    cho_phep_xem = models.BooleanField(
        default=True, 
        verbose_name="Cho phép xem giấy báo"
    )
    ngay_bat_dau = models.DateTimeField(
        null=True, 
        blank=True, 
        verbose_name="Ngày bắt đầu tra cứu"
    )
    ngay_ket_thuc = models.DateTimeField(
        null=True, 
        blank=True, 
        verbose_name="Ngày kết thúc tra cứu"
    )

    class Meta:
        db_table = 'cau_hinh_giay_bao'

    @classmethod
    def get_config(cls):
        config, _ = cls.objects.get_or_create(id=1)
        return config

    @property
    def is_open(self):
        # 1. Ưu tiên kiểm tra công tắc thủ công
        if not self.cho_phep_xem:
            return False

        # 2. Kiểm tra theo khoảng thời gian đặt lịch
        now = timezone.now()
        if self.ngay_bat_dau and now < self.ngay_bat_dau:
            return False
        if self.ngay_ket_thuc and now > self.ngay_ket_thuc:
            return False

        return True

    @property
    def status_text(self):
        # 1. Nếu tắt công tắc thủ công
        if not self.cho_phep_xem:
            return "Tính năng xem giấy báo đang tạm khóa"

        # 2. Kiểm tra các mốc thời gian (quy đổi timezone về giờ hệ thống)
        now = timezone.now()
        if self.ngay_bat_dau and now < self.ngay_bat_dau:
            dt_local = timezone.localtime(self.ngay_bat_dau)
            return f"Chưa đến hạn mở tra cứu (Mở lúc {dt_local.strftime('%H:%M %d/%m/%Y')})"
        
        if self.ngay_ket_thuc and now > self.ngay_ket_thuc:
            return "Đã hết thời hạn tra cứu / Đã khóa"

        return "Đã sẵn sàng tải bản chính thức"
