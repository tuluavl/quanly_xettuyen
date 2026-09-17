from functools import wraps
from django.shortcuts import redirect
from django.contrib import messages

def custom_login_required(view_func):
    """
    Decorator bắt buộc đăng nhập. 
    Nếu hết hạn session (idle timeout), hiển thị thông báo và yêu cầu đăng nhập lại.
    """
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.session.get('user_id'):
            messages.warning(request, 'Phiên làm việc đã hết hạn hoặc bạn chưa đăng nhập!')
            return redirect('login')
        return view_func(request, *args, **kwargs)
    return _wrapped_view


def check_permission(perm_code):
    """
    Decorator kiểm tra quyền truy cập của người dùng.
    """
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            # 1. Chưa đăng nhập -> Chuyển hướng kèm thông báo
            if not request.session.get('user_id'):
                messages.warning(request, 'Phiên làm việc đã hết hạn. Vui lòng đăng nhập lại!')
                return redirect('login')
            
            # 2. TOÀN QUYỀN CHO ADMIN (Kiểm tra theo role_code, role_name và username)
            role_code = str(request.session.get('role_code', '')).lower().strip()
            role_name = str(request.session.get('role_name', '')).lower().strip()
            username = str(request.session.get('username', '')).lower().strip()
            
            is_admin = (
                role_code == 'admin' or 
                username == 'admin' or 
                role_name in ['quản trị viên', 'administrator']
            )
            
            if is_admin:
                return view_func(request, *args, **kwargs)

            # 3. Kiểm tra danh sách quyền của các vai trò khác (Cán bộ,...)
            user_perms = request.session.get('permissions', [])
            if perm_code not in user_perms:
                messages.error(request, f'Bạn không có quyền thực hiện chức năng này ({perm_code})!')
                return redirect('index')
                
            return view_func(request, *args, **kwargs)
        return _wrapped_view
    return decorator