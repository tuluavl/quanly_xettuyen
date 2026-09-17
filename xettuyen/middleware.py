from django.contrib.auth import logout
from django.shortcuts import redirect
from django.contrib import messages
from .models import CustomUser

class CheckUserStatusMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        user_id = request.session.get('user_id')

        if user_id:
            user = CustomUser.objects.filter(id=user_id).select_related('role').first()

            # 1. Kiểm tra tài khoản không tồn tại hoặc bị khóa
            is_locked = (
                not user or
                (hasattr(user, 'is_active') and not user.is_active) or
                getattr(user, 'is_locked', False) is True or
                getattr(user, 'trang_thai', None) in ['locked', 'inactive', 0, False, 'khoa'] or
                getattr(user, 'status', None) in ['locked', 'inactive', 0, False, 'khoa']
            )

            if is_locked:
                request.session.flush()  # Cưỡng chế hủy Session ngay lập tức
                messages.error(request, 'Tài khoản của bạn đã bị khóa! Vui lòng liên hệ quản trị viên.')
                return redirect('login')

            # 2. Cập nhật trực tiếp quyền hạn mới nhất nếu Admin có thay đổi trên giao diện
            if user.role:
                request.session['role_id'] = user.role.id
                request.session['role_name'] = user.role.name
                request.session['role_code'] = user.role.code
                request.session['permissions'] = list(user.role.permissions.values_list('code', flat=True))
            else:
                request.session['role_id'] = None
                request.session['role_name'] = ''
                request.session['role_code'] = ''
                request.session['permissions'] = []

        return self.get_response(request)