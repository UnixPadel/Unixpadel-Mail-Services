"""
UnixPadel Mail Services - Auto Updater Module
Provides background version checking, silent background downloading,
and Windows self-replacement & restart logic.
"""

import os
import sys
import json
import time
import urllib.request
import subprocess
import threading

APP_VERSION = "5.2.1"
GITHUB_REPO = "UnixPadel/Unixpadel-Mail-Services"
RELEASE_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

def parse_version(v_str: str):
    """
    Converts version strings like 'v5.2.1' or '5.2.0' to a tuple of integers for comparison.
    """
    if not v_str:
        return (0, 0, 0)
    cleaned = v_str.strip().lstrip('vV')
    parts = []
    for p in cleaned.split('.'):
        digits = ''.join(c for c in p if c.isdigit())
        if digits:
            parts.append(int(digits))
        else:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts)


def get_latest_release_info(token: str = None):
    """
    Fetches the latest release info from GitHub.
    Returns (latest_version_str, download_url) or (None, None) if check fails or no asset found.
    """
    try:
        headers = {
            "User-Agent": "UnixPadel-Mail-Services-Updater",
            "Accept": "application/vnd.github+json"
        }
        if token:
            headers["Authorization"] = f"Bearer {token}"

        req = urllib.request.Request(RELEASE_API_URL, headers=headers)
        with urllib.request.urlopen(req, timeout=10) as response:
            if response.status == 200:
                data = json.loads(response.read().decode('utf-8'))
                tag_name = data.get("tag_name", "")
                assets = data.get("assets", [])

                exe_url = None
                for asset in assets:
                    name = asset.get("name", "").lower()
                    if name.endswith(".exe"):
                        exe_url = asset.get("browser_download_url")
                        break

                return tag_name, exe_url
    except Exception as e:
        print(f"[Updater] Version check failed: {e}")
    
    return None, None


def download_file(url: str, dest_path: str, progress_callback=None):
    """
    Downloads a file from url to dest_path with progress callback.
    """
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "UnixPadel-Mail-Services-Updater"})
        with urllib.request.urlopen(req, timeout=60) as response:
            total_size = int(response.headers.get('Content-Length', 0))
            downloaded = 0
            block_size = 65536 # 64KB chunks

            with open(dest_path, 'wb') as f:
                while True:
                    buffer = response.read(block_size)
                    if not buffer:
                        break
                    downloaded += len(buffer)
                    f.write(buffer)
                    if progress_callback and total_size > 0:
                        pct = int((downloaded / total_size) * 100)
                        progress_callback(pct)
        return True
    except Exception as e:
        print(f"[Updater] Download failed: {e}")
        if os.path.exists(dest_path):
            try:
                os.remove(dest_path)
            except Exception:
                pass
        return False


def apply_update_and_restart(new_exe_path: str, app_instance=None):
    """
    Launches a Windows batch script to replace current running executable and restart.
    """
    if not getattr(sys, 'frozen', False):
        print(f"[Updater] Running from source mode. New executable saved to: {new_exe_path}")
        return

    current_exe = sys.executable
    current_pid = os.getpid()
    temp_dir = os.environ.get("TEMP", os.path.dirname(current_exe))
    bat_path = os.path.join(temp_dir, "unixpadel_update_helper.bat")

    bat_content = f"""@echo off
chcp 65001 >nul
echo UnixPadel Mail Services guncelleniyor, lutfen bekleyin...
set TARGET="{current_exe}"
set NEW_EXE="{new_exe_path}"
set PID={current_pid}

:wait_loop
tasklist /fi "PID eq %PID%" 2>nul | find "%PID%" >nul
if errorlevel 1 goto do_update
timeout /t 1 /nobreak >nul
goto wait_loop

:do_update
timeout /t 1 /nobreak >nul
copy /y %NEW_EXE% %TARGET% >nul
del %NEW_EXE% >nul
start "" %TARGET%
del "%~f0"
"""

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    print("[Updater] Launching update helper and restarting application...")
    
    # Run batch script detached
    CREATE_NO_WINDOW = 0x08000000
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(
        [bat_path],
        creationflags=CREATE_NO_WINDOW | DETACHED_PROCESS,
        close_fds=True
    )

    if app_instance:
        try:
            app_instance.destroy()
        except Exception:
            pass

    os._exit(0)


def start_silent_auto_update(status_callback=None, app_instance=None):
    """
    Runs the full update check and silent installation in a background thread.
    """
    def _worker():
        # Short initial delay to let app UI initialize completely
        time.sleep(3)
        
        remote_tag, exe_url = get_latest_release_info()
        if not remote_tag or not exe_url:
            return

        current_v = parse_version(APP_VERSION)
        remote_v = parse_version(remote_tag)

        if remote_v > current_v:
            print(f"[Updater] New update found: {remote_tag} (Current: v{APP_VERSION})")
            if status_callback:
                status_callback(f"v{APP_VERSION} (Yeni Sürüm İndiriliyor...)")

            temp_dir = os.environ.get("TEMP", ".")
            temp_exe = os.path.join(temp_dir, f"UnixPadel_Update_{remote_tag}.exe")

            def _on_progress(pct):
                if status_callback:
                    status_callback(f"v{APP_VERSION} (İndiriliyor %{pct})")

            success = download_file(exe_url, temp_exe, progress_callback=_on_progress)
            if success and os.path.exists(temp_exe):
                if status_callback:
                    status_callback(f"v{APP_VERSION} (Yeniden Başlatılıyor...)")
                time.sleep(1)
                apply_update_and_restart(temp_exe, app_instance=app_instance)
            else:
                if status_callback:
                    status_callback(f"v{APP_VERSION}")

    thread = threading.Thread(target=_worker, daemon=True)
    thread.start()
