"""
launcher.py
-----------
PyInstaller로 빌드된 EXE의 진입점.
Streamlit 서버를 subprocess로 시작하고 브라우저를 자동으로 엽니다.
"""
import sys
import os
import time
import socket
import threading
import subprocess
import webbrowser


# ──────────────────────────────────────────────
# 1. 빌드 환경에서 리소스 경로 확인
# ──────────────────────────────────────────────
def resource_path(relative: str) -> str:
    """PyInstaller 패키징 후 실제 파일 경로 반환."""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)


# ──────────────────────────────────────────────
# 2. 빈 포트 자동 탐색
# ──────────────────────────────────────────────
def find_free_port(start: int = 8501) -> int:
    for port in range(start, start + 100):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return start


# ──────────────────────────────────────────────
# 3. Streamlit 서버 프로세스 시작
# ──────────────────────────────────────────────
def start_streamlit(app_path: str, port: int) -> subprocess.Popen:
    # PyInstaller 번들 내부의 streamlit 실행 파일 우선 탐색
    if hasattr(sys, "_MEIPASS"):
        st_exe = os.path.join(sys._MEIPASS, "streamlit.exe")
        if not os.path.exists(st_exe):
            st_exe = os.path.join(sys._MEIPASS, "streamlit", "__main__.py")
    else:
        st_exe = None

    env = os.environ.copy()
    # 현재 exe(또는 스크립트) 디렉터리를 PYTHONPATH에 추가
    exe_dir = os.path.dirname(os.path.abspath(sys.executable if hasattr(sys, "_MEIPASS") else __file__))
    env["PYTHONPATH"] = exe_dir + os.pathsep + env.get("PYTHONPATH", "")

    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.enableCORS", "false",
        "--server.enableXsrfProtection", "false",
        "--browser.gatherUsageStats", "false",
        "--theme.base", "dark",
    ]

    return subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
        # 터미널 창 숨기기 (Windows)
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )


# ──────────────────────────────────────────────
# 4. 서버가 LISTEN 상태가 될 때까지 대기
# ──────────────────────────────────────────────
def wait_for_server(port: int, timeout: int = 30) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except OSError:
            time.sleep(0.3)
    return False


# ──────────────────────────────────────────────
# 5. 메인
# ──────────────────────────────────────────────
def main():
    app_path = resource_path("app.py")
    port = find_free_port(8501)
    url  = f"http://127.0.0.1:{port}"

    proc = start_streamlit(app_path, port)

    if wait_for_server(port, timeout=40):
        webbrowser.open(url)
        # 프로세스 종료 시까지 대기 (EXE가 살아있는 한 서버도 유지)
        proc.wait()
    else:
        print("[ERROR] Streamlit 서버가 시작되지 않았습니다.")
        proc.terminate()
        sys.exit(1)


if __name__ == "__main__":
    main()
