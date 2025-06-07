"""
mini_updater.py – self-update helper for GA Broker Helper

• Contacts the latest-release endpoint on GitHub
• Looks for    GA_broker_helper_<tag>.zip
• Extracts it next to the current folder
• Runs the new EXE, schedules the old folder for deletion with a .bat
"""

from __future__ import annotations
import os, shutil, subprocess, sys, tempfile, zipfile, textwrap
from pathlib import Path
import requests

# ---------------------------------------------------------------------------
__version__ = "1.6.0"  # 🔁 bump before each release
REPO_API    = "https://api.github.com/repos/Marshall-Ye/Broker-Helper/releases/latest"
ASSET_PREFIX = "GA_broker_helper_"
EXE_NAME     = "GA Broker Helper.exe"
TIMEOUT      = 10
# ---------------------------------------------------------------------------


def _latest_release() -> tuple[str, str]:
    data = requests.get(REPO_API, timeout=TIMEOUT).json()
    tag  = data["tag_name"].lstrip("v")
    url  = next(a["browser_download_url"]
                for a in data["assets"]
                if a["name"] == f"{ASSET_PREFIX}{tag}.zip")
    return tag, url


def _download(url: str, dest: Path) -> None:
    with requests.get(url, stream=True, timeout=30) as r, dest.open("wb") as f:
        for chunk in r.iter_content(65536):
            f.write(chunk)


def check_and_update() -> str:
    try:
        tag, url = _latest_release()
        if tag == __version__:
            return "latest"

        tmp = Path(tempfile.mkdtemp())
        zip_path = tmp / "update.zip"
        _download(url, zip_path)

        run_dir  = Path(sys.executable).resolve().parent  # …\GA_broker_helper_<old>
        root_dir = run_dir.parent
        new_dir  = root_dir / f"{ASSET_PREFIX}{tag}"

        if new_dir.exists():
            shutil.rmtree(new_dir, ignore_errors=True)

        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(new_dir)

        new_exe = new_dir / EXE_NAME
        if not new_exe.exists():
            raise FileNotFoundError(f"{EXE_NAME} missing in the update zip")

        # --- cleanup script to delete the old folder
        bat = tmp / "cleanup.bat"
        bat.write_text(textwrap.dedent(f"""\

            @echo off
            cd /d %TEMP%
            timeout /t 6 >nul
            rmdir /s /q "{run_dir}"

            echo Update successful. Old version removed.
            echo You can close this window now.
            echo.
            start "" /min cmd /c "ping 127.0.0.1 -n 3 >nul & del \"%~f0\""
            pause >nul

        """))

        subprocess.Popen(
            ["cmd", "/c", "start", "", str(bat)],
            cwd=str(tmp),
            creationflags=subprocess.CREATE_NO_WINDOW,
            close_fds=True
        )

        # --- launch new version
        subprocess.Popen([str(new_exe)], cwd=str(new_dir), close_fds=True)

        # change working dir and exit
        os.chdir(str(Path.home()))
        os._exit(0)

    except Exception as e:
        return f"error:{e}"
