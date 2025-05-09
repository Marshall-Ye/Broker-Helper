"""
mini_updater.py  –  GA Broker Helper self-update helper
────────────────────────────────────────────────────────
• Bump __version__ before each release
• In main_ui.py do:  VERSION = updater.__version__
• Call updater.check_and_update() to pull down and install
"""

from __future__ import annotations
import os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

import requests
import pythoncom
from win32com.client import Dispatch

# ───────────────────────────────────────────────────────────
__version__ = "1.4.3"                         # ← bump this
REPO_API    = (
    "https://api.github.com/repos/Marshall-Ye/"
    "Broker-Helper/releases/latest"
)
# ───────────────────────────────────────────────────────────


# ╭─────────────────── helper functions ───────────────────╮
def _latest_release() -> tuple[str, str]:
    """
    Return (tag, download_url) for the asset
    GA_broker_helper_{tag}.zip in the latest GitHub release.
    """
    data = requests.get(REPO_API, timeout=10).json()
    tag  = data["tag_name"].lstrip("v")
    expected = f"GA_broker_helper_{tag}.zip"
    asset = next(a for a in data["assets"] if a["name"] == expected)
    return tag, asset["browser_download_url"]


def _download(url: str, dest: os.PathLike) -> None:
    """Stream-download *url* into local file *dest*."""
    with requests.get(url, stream=True, timeout=30) as r, open(dest, "wb") as f:
        for chunk in r.iter_content(65536):
            f.write(chunk)


def _prune_old_versions(root: Path, keep: int = 2) -> None:
    """Keep only the *keep* most recent GA_broker_helper_* folders."""
    versions: list[tuple[tuple[int, ...], Path]] = []
    for entry in root.iterdir():
        if entry.is_dir():
            m = re.fullmatch(r"GA_broker_helper_(\d+\.\d+\.\d+)", entry.name)
            if m:
                vt = tuple(map(int, m.group(1).split(".")))
                versions.append((vt, entry))
    versions.sort(reverse=True)          # newest first
    for _, path in versions[keep:]:
        shutil.rmtree(path, ignore_errors=True)


def _update_shortcut(folder: Path, target_exe: Path) -> None:
    """Create or refresh GA Broker Helper.lnk inside *folder*."""
    pythoncom.CoInitialize()
    link_path = folder / "GA Broker Helper.lnk"
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(str(link_path))
    shortcut.Targetpath       = str(target_exe)
    shortcut.WorkingDirectory = str(target_exe.parent)
    shortcut.IconLocation     = str(target_exe)
    shortcut.save()
    pythoncom.CoUninitialize()


# ╭──────────────────── main entry point ───────────────────╮
# ... all earlier helper functions unchanged ...

def check_and_update() -> str:
    try:
        latest, url = _latest_release()
        if latest == __version__:
            return "latest"

        tmp_dir  = Path(tempfile.mkdtemp())
        zip_path = tmp_dir / "u.zip"
        _download(url, zip_path)

        # our exe lives in …\<old>\_internal\GA Broker Helper.exe
        exe_dir      = Path(sys.executable).resolve().parent          # _internal
        install_root = exe_dir.parent                                 # <old>

        target_dir = install_root / f"GA_broker_helper_{latest}"
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(target_dir)

        # launch the level-1 exe
        new_exe = target_dir / "_internal" / "GA Broker Helper.exe"   # ← changed
        if not new_exe.exists():
            raise FileNotFoundError("GA Broker Helper.exe missing")

        _prune_old_versions(install_root, keep=2)
        _update_shortcut(install_root, new_exe)
        subprocess.Popen([str(new_exe)], close_fds=True)
        os._exit(0)

    except Exception as e:
        return f"error:{e}"
