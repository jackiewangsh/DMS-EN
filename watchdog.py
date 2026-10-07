# DMS 守护脚本 — 检测服务是否存活，崩溃则自动重启
# 由 Windows 计划任务每60秒调用一次
import urllib.request
import urllib.error
import subprocess
import sys
import os
import time

HEALTH_URL = "http://localhost:5000/api/health"
MAX_RETRIES = 2
RETRY_DELAY = 5  # seconds

def is_healthy():
    """检查 DMS 健康状态"""
    for attempt in range(MAX_RETRIES):
        try:
            resp = urllib.request.urlopen(HEALTH_URL, timeout=10)
            if resp.status == 200:
                return True
        except (urllib.error.URLError, ConnectionError, OSError):
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
    return False

def restart_dms():
    """杀掉旧进程并重启 DMS 服务"""
    # 杀掉占用5000端口的进程
    try:
        result = subprocess.run(
            ['netstat', '-ano'],
            capture_output=True, text=True, timeout=10
        )
        for line in result.stdout.split('\n'):
            if ':5000' in line and 'LISTENING' in line:
                parts = line.strip().split()
                pid = parts[-1]
                if pid.isdigit():
                    subprocess.run(['taskkill', '/F', '/PID', pid], timeout=10)
                    print(f"[DMS-Watchdog] Killed stale process PID={pid}")
    except Exception as e:
        print(f"[DMS-Watchdog] Cleanup error: {e}")

    # 启动 DMS（使用 _run.py + Waitress）
    dms_dir = os.path.dirname(os.path.abspath(__file__))
    python_exe = os.path.join(dms_dir, 'venv', 'Scripts', 'python.exe')
    run_server = os.path.join(dms_dir, '_run.py')
    
    if not os.path.exists(python_exe):
        print(f"[DMS-Watchdog] Python not found: {python_exe}")
        return False
    
    try:
        # 使用 CREATE_NEW_PROCESS_GROUP 使进程独立于当前控制台
        subprocess.Popen(
            [python_exe, run_server],
            cwd=dms_dir,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.DETACHED_PROCESS,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
        print(f"[DMS-Watchdog] DMS restarted at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        return True
    except Exception as e:
        print(f"[DMS-Watchdog] Restart failed: {e}")
        return False

if __name__ == '__main__':
    if is_healthy():
        # 静默退出 — 一切正常
        sys.exit(0)
    else:
        print(f"[DMS-Watchdog] {time.strftime('%Y-%m-%d %H:%M:%S')} - DMS unhealthy, restarting...")
        restart_dms()
