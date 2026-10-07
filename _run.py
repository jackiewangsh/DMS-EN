"""启动 DMS 服务，监听 0.0.0.0:5000（生产模式）"""
import os
import sys
import gc
import signal
import threading
import traceback

# 确保工作目录正确
os.chdir(os.path.dirname(os.path.abspath(__file__)))

from app import app
from waitress import serve

# 定期内存清理（每30分钟执行一次，防止长时间运行的内存泄漏）
def _periodic_gc():
    gc.collect()
    timer = threading.Timer(1800, _periodic_gc)
    timer.daemon = True
    timer.start()

# 优雅关闭
def _signal_handler(signum, frame):
    print(f'\n[DMS] 收到信号 {signum}，正在关闭...')
    sys.exit(0)

signal.signal(signal.SIGINT, _signal_handler)
signal.signal(signal.SIGTERM, _signal_handler)

if __name__ == '__main__':
    print('=' * 50)
    print('  DMS Production Server (Waitress)')
    print('  http://0.0.0.0:5000')
    print('=' * 50)

    # 启动定期内存清理
    _periodic_gc()

    serve(
        app,
        host='0.0.0.0',
        port=5000,
        threads=16,              # 15人在线，每人2-3并发请求，峰值约45
        channel_timeout=120,      # 长请求超时（LibreOffice转换等）
        cleanup_interval=30,      # 清理闲置连接间隔（秒）
        connection_limit=200,     # 最大连接数（含静态资源）
        max_request_body_size=500 * 1024 * 1024,  # 500MB
    )
