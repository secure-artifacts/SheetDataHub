from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
import sys
import tempfile
import winreg
from pathlib import Path

from sheet_hub.version import APP_VERSION

APP_NAME = "表数通"
APP_ID = "SheetDataHub"
VERSION = APP_VERSION
EXE_NAME = "SheetDataHub.exe"
MB_OK, MB_YESNO = 0, 4
MB_ICONINFORMATION, MB_ICONQUESTION, MB_ICONERROR = 0x40, 0x20, 0x10
IDYES = 6


def message(text: str, title: str, flags: int = MB_OK) -> int:
    return ctypes.windll.user32.MessageBoxW(None, text, title, flags)


def bundled(relative: str) -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / relative


def install_dir() -> Path:
    return Path(os.environ["LOCALAPPDATA"]) / "Programs" / APP_ID


def shortcut(path: Path, target: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    escaped_path = str(path).replace("'", "''")
    escaped_target = str(target).replace("'", "''")
    escaped_working = str(target.parent).replace("'", "''")
    script = (
        "$w=New-Object -ComObject WScript.Shell;"
        f"$s=$w.CreateShortcut('{escaped_path}');"
        f"$s.TargetPath='{escaped_target}';"
        f"$s.WorkingDirectory='{escaped_working}';"
        f"$s.IconLocation='{escaped_target}';$s.Save()"
    )
    subprocess.run(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
        check=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def register_uninstall(folder: Path) -> None:
    key_path = rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_ID}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        values = {
            "DisplayName": APP_NAME,
            "DisplayVersion": VERSION,
            "Publisher": APP_ID,
            "DisplayIcon": str(folder / EXE_NAME),
            "InstallLocation": str(folder),
            "UninstallString": f'"{folder / "Uninstall.exe"}" --uninstall',
        }
        for name, value in values.items():
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        winreg.SetValueEx(key, "NoModify", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "NoRepair", 0, winreg.REG_DWORD, 1)


def shortcut_paths() -> tuple[Path, Path]:
    desktop = Path(os.environ.get("USERPROFILE", "")) / "Desktop" / f"{APP_NAME}.lnk"
    start = Path(os.environ["APPDATA"]) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / f"{APP_NAME}.lnk"
    return desktop, start


def uninstall() -> None:
    if message("确定卸载表数通吗？\n\n个人配置、汇总库和日志将会保留。", "卸载表数通", MB_YESNO | MB_ICONQUESTION) != IDYES:
        return
    for path in shortcut_paths():
        if path.exists():
            path.unlink()
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rf"Software\Microsoft\Windows\CurrentVersion\Uninstall\{APP_ID}")
    except FileNotFoundError:
        pass
    folder = install_dir()
    command = f'ping 127.0.0.1 -n 3 >nul & rmdir /s /q "{folder}"'
    subprocess.Popen(["cmd.exe", "/c", command], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    message("表数通已卸载。个人数据仍保存在本地应用数据目录。", "卸载完成", MB_ICONINFORMATION)


def install() -> None:
    folder = install_dir()
    prompt = f"安装 {APP_NAME} v{VERSION}？\n\n安装位置：\n{folder}"
    if message(prompt, f"安装 {APP_NAME}", MB_YESNO | MB_ICONQUESTION) != IDYES:
        return
    try:
        folder.mkdir(parents=True, exist_ok=True)
        target = folder / EXE_NAME
        shutil.copytree(bundled("payload/SheetDataHub"), folder, dirs_exist_ok=True)
        shutil.copy2(sys.executable, folder / "Uninstall.exe")
        desktop, start = shortcut_paths()
        shortcut(start, target)
        if message("是否创建桌面快捷方式？", APP_NAME, MB_YESNO | MB_ICONQUESTION) == IDYES:
            shortcut(desktop, target)
        register_uninstall(folder)
        if message(f"{APP_NAME} 已成功安装。\n\n是否现在启动？", "安装完成", MB_YESNO | MB_ICONINFORMATION) == IDYES:
            subprocess.Popen([str(target)], cwd=str(folder))
    except Exception as exc:
        message(f"安装失败：\n{exc}", "安装失败", MB_ICONERROR)
        raise


def verify_payload() -> int:
    with tempfile.TemporaryDirectory(prefix="sheet-data-hub-test-") as data_dir:
        environment = os.environ.copy()
        environment["SHEET_DATA_HUB_DATA_DIR"] = data_dir
        completed = subprocess.run(
            [str(bundled("payload/SheetDataHub/SheetDataHub.exe")), "--smoke-test"],
            env=environment,
            timeout=30,
        )
        return completed.returncode


if __name__ == "__main__":
    if "--verify-payload" in sys.argv:
        raise SystemExit(verify_payload())
    uninstall() if "--uninstall" in sys.argv else install()
