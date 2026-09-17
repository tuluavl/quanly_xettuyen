from waitress import serve
from quanly_tuyensinh.wsgi import application

if __name__ == '__main__':
    print("Server đang chạy tại http://0.0.0.0:8000")
    serve(application, host='0.0.0.0', port=8000)