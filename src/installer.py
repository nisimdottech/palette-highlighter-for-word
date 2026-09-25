from __future__ import annotations

import hashlib
import os
import re
import shutil
import subprocess
import sys
import time
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import messagebox


APP_NAME = "Palette Highlighter for Word"
APP_ID = "PaletteHighlighterForWord"
APP_PUBLISHER = "Nisim Levi"
APP_VERSION = "1.0.0"
UNINSTALL_ROOT = r"Software\Microsoft\Windows\CurrentVersion\Uninstall"
REG_PATH = rf"{UNINSTALL_ROOT}\{APP_ID}"
# Name used before the rename; its template would load a second Highlighter tab.
LEGACY_ID = "CustomHighlighter"
BACKUPS_KEPT = 3
BACKUP_PATTERN = re.compile(r"_(\d{4}-\d{2}-\d{2}_\d{6})\.dotm\.bak$")


def startup_dir() -> Path:
    return Path(os.environ["APPDATA"]) / "Microsoft" / "Word" / "STARTUP"


def startup_path() -> Path:
    return startup_dir() / f"{APP_ID}.dotm"


def launcher_dir(app_id: str = APP_ID) -> Path:
    return Path(os.environ.get("LOCALAPPDATA", startup_dir())) / app_id


def backup_dir() -> Path:
    return launcher_dir() / "backups"


def bundled_template() -> Path:
    root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return root / f"{APP_ID}.dotm"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def word_running() -> bool:
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/NH"],
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        check=False,
    )
    return "WINWORD.EXE" in result.stdout.upper()


def wait_for_word_to_close(root: tk.Tk) -> bool:
    while word_running():
        retry = messagebox.askretrycancel(
            APP_NAME,
            "Microsoft Word is running. Close Word before installing or uninstalling, then choose Retry.",
            parent=root,
        )
        if not retry:
            return False
        time.sleep(0.25)
    return True


def backup_stamp(backup: Path) -> str:
    match = BACKUP_PATTERN.search(backup.name)
    return match.group(1) if match else ""


def prune_backups() -> None:
    backups = sorted(backup_dir().glob("*.dotm.bak"), key=backup_stamp, reverse=True)
    for old_backup in backups[BACKUPS_KEPT:]:
        old_backup.unlink()


def move_startup_backups() -> None:
    # Earlier versions left their backups in STARTUP.
    backup_dir().mkdir(parents=True, exist_ok=True)
    for app_id in (APP_ID, LEGACY_ID):
        for backup in startup_dir().glob(f"{app_id}_*.dotm.bak"):
            if BACKUP_PATTERN.search(backup.name):
                shutil.move(backup, backup_dir() / backup.name)
    prune_backups()


def backup_existing(target: Path) -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    backup_dir().mkdir(parents=True, exist_ok=True)
    backup = backup_dir() / f"{target.stem}_{stamp}.dotm.bak"
    shutil.copy2(target, backup)
    prune_backups()
    return backup


def save_uninstall_record(target: Path, installed_hash: str) -> None:
    import winreg

    launcher = launcher_dir() / f"{APP_ID}.exe"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
        values = {
            "DisplayName": APP_NAME,
            "DisplayVersion": APP_VERSION,
            "Publisher": APP_PUBLISHER,
            "InstallLocation": str(target.parent),
            "UninstallString": f'"{launcher}" --uninstall',
            "InstalledTemplateSha256": installed_hash,
        }
        for name, value in values.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)


def read_uninstall_hash() -> str | None:
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH) as key:
            return winreg.QueryValueEx(key, "InstalledTemplateSha256")[0]
    except FileNotFoundError:
        return None


def remove_uninstall_record() -> None:
    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REG_PATH)
    except FileNotFoundError:
        pass


def remove_legacy_install() -> None:
    import winreg

    legacy = startup_dir() / f"{LEGACY_ID}.dotm"
    if legacy.exists():
        backup_existing(legacy)
        legacy.unlink()
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"{UNINSTALL_ROOT}\{LEGACY_ID}")
    except FileNotFoundError:
        pass
    shutil.rmtree(launcher_dir(LEGACY_ID), ignore_errors=True)


def install(root: tk.Tk) -> bool:
    template = bundled_template()
    target = startup_path()
    if not template.is_file():
        messagebox.showerror(APP_NAME, f"The bundled {APP_ID}.dotm is missing.", parent=root)
        return False
    if not wait_for_word_to_close(root):
        return False
    if target.exists():
        replace = messagebox.askyesno(
            APP_NAME,
            f"{APP_NAME} is already installed.\n\nDo you want to replace the existing version?",
            parent=root,
        )
        if not replace:
            return False
        try:
            backup_existing(target)
        except OSError as error:
            messagebox.showerror(APP_NAME, f"Could not back up the existing template:\n{error}", parent=root)
            return False
    try:
        remove_legacy_install()
        move_startup_backups()
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(template, target)
        installed_hash = sha256(target)
        if installed_hash != sha256(template):
            raise OSError("The installed template failed verification.")
        launcher_dir().mkdir(parents=True, exist_ok=True)
        launcher = launcher_dir() / f"{APP_ID}.exe"
        if Path(sys.executable).resolve() != launcher.resolve():
            shutil.copy2(sys.executable, launcher)
        save_uninstall_record(target, installed_hash)
    except OSError as error:
        messagebox.showerror(APP_NAME, f"Installation failed:\n{error}", parent=root)
        return False
    messagebox.showinfo(
        APP_NAME,
        f"{APP_NAME} has been installed.\n\nRestart Microsoft Word to load the Highlighter tab.",
        parent=root,
    )
    return True


def uninstall(root: tk.Tk) -> bool:
    target = startup_path()
    if not wait_for_word_to_close(root):
        return False
    expected = read_uninstall_hash()
    if target.exists() and expected and sha256(target) == expected:
        try:
            target.unlink()
        except OSError as error:
            messagebox.showerror(APP_NAME, f"Could not remove the installed template:\n{error}", parent=root)
            return False
    elif target.exists():
        messagebox.showwarning(
            APP_NAME,
            "The STARTUP template has changed since installation, so it was left untouched.",
            parent=root,
        )
    remove_uninstall_record()
    messagebox.showinfo(APP_NAME, f"{APP_NAME} was uninstalled.", parent=root)
    return True


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    try:
        if "--uninstall" in sys.argv[1:]:
            uninstall(root)
        else:
            install(root)
    finally:
        root.destroy()


if __name__ == "__main__":
    main()
