"""
mini_updater.py – self-update helper that deletes the old folder
via an external .bat script.
"""
from __future__ import annotations
import os, shutil, subprocess, sys, tempfile, zipfile, textwrap
from pathlib import Path
import requests

__version__ = "1.4.7"      # bump each release
REPO_API    = ("https://api.github.com/repos/Marshall-Ye/"
               "Broker-Helper/releases/latest")


def _latest_release() -> tuple[str, str]:
    data = requests.get(REPO_API, timeout=10).json()
    tag  = data["tag_name"].lstrip("v")
    url  = next(a["browser_download_url"]
                for a in data["assets"]
                if a["name"] == f"GA_broker_helper_{tag}.zip")
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
        zip_path = tmp / "upd.zip"
        _download(url, zip_path)

        run_dir  = Path(sys.executable).resolve().parent          # …\GA_broker_helper_<old>
        root_dir = run_dir.parent
        new_dir  = root_dir / f"GA_broker_helper_{tag}"

        if new_dir.exists():
            shutil.rmtree(new_dir, ignore_errors=True)
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(new_dir)

        new_exe = new_dir / "GA Broker Helper.exe"
        if not new_exe.exists():
            raise FileNotFoundError("Update zip missing GA Broker Helper.exe")

            # … earlier code unchanged …

            # 4) write & launch cleanup .bat --------------------------
            bat = tmp / "cleanup.bat"
            bat.write_text(textwrap.dedent(f"""\
                    @echo off
                    ping 127.0.0.1 -n 6 >nul
                    rmdir /s /q "{run_dir}"
                    del "%~f0"
                """))
            subprocess.Popen(
                ["cmd", "/c", "start", "", str(bat)],
                creationflags=subprocess.CREATE_NO_WINDOW,
                close_fds=True
            )

            # 5) launch the new build --------------------------------
            subprocess.Popen([str(new_exe)], cwd=str(new_dir), close_fds=True)

            os._exit(0)


    except Exception as e:
        return f"error:{e}"
