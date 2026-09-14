from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import requests


APP_VERSION = "1.2.8"
UPDATE_REPO = "secure-artifacts/SheetDataHub"
RELEASES_URL = f"https://github.com/{UPDATE_REPO}/releases"


def parse_version(text: str) -> tuple[int, ...]:
    cleaned = str(text or "").strip().lstrip("vV")
    parts: list[int] = []
    for item in re.split(r"[.\-+_]", cleaned):
        if item.isdigit():
            parts.append(int(item))
        elif parts:
            break
    return tuple(parts or (0,))


def is_newer(latest: str, current: str = APP_VERSION) -> bool:
    left, right = parse_version(latest), parse_version(current)
    width = max(len(left), len(right))
    return left + (0,) * (width - len(left)) > right + (0,) * (width - len(right))


def fetch_latest_release(timeout: int = 20) -> dict[str, Any]:
    response = requests.get(
        f"https://api.github.com/repos/{UPDATE_REPO}/releases/latest",
        timeout=timeout,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "SheetDataHub",
        },
    )
    if response.status_code == 404:
        raise RuntimeError("还没有发布版本，稍后再试。")
    response.raise_for_status()
    data = response.json()
    tag = str(data.get("tag_name") or "").strip()
    version = tag.lstrip("vV") or APP_VERSION
    assets = list(data.get("assets") or [])
    installer = next(
        (
            asset for asset in assets
            if str(asset.get("name") or "").casefold() == "sheetdatahub-setup.exe"
        ),
        None,
    )
    return {
        "tag": tag or f"v{version}",
        "version": version,
        "url": str(data.get("html_url") or RELEASES_URL),
        "name": str(data.get("name") or tag or version),
        "installer_url": str((installer or {}).get("browser_download_url") or ""),
        "installer_size": int((installer or {}).get("size") or 0),
    }


def download_release_installer(
    info: dict[str, Any],
    data_dir: str | Path,
    timeout: int = 180,
) -> Path:
    url = str(info.get("installer_url") or "").strip()
    if not url:
        raise RuntimeError("该版本没有可用的 Windows 安装包，请稍后再试。")
    version = re.sub(r"[^0-9A-Za-z._-]", "_", str(info.get("version") or "latest"))
    update_dir = Path(data_dir) / "updates"
    update_dir.mkdir(parents=True, exist_ok=True)
    destination = update_dir / f"SheetDataHub-Setup-v{version}.exe"
    partial = destination.with_suffix(".exe.part")
    expected_size = int(info.get("installer_size") or 0)
    try:
        with requests.get(
            url,
            stream=True,
            timeout=(20, timeout),
            headers={"User-Agent": "SheetDataHub"},
        ) as response:
            response.raise_for_status()
            written = 0
            with partial.open("wb") as output:
                for chunk in response.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        output.write(chunk)
                        written += len(chunk)
        if expected_size and written != expected_size:
            raise RuntimeError(f"安装包下载不完整：应为 {expected_size} 字节，实际 {written} 字节。")
        if written < 1024 * 1024:
            raise RuntimeError("下载到的安装包大小异常。")
        with partial.open("rb") as source:
            if source.read(2) != b"MZ":
                raise RuntimeError("下载内容不是有效的 Windows 安装程序。")
        partial.replace(destination)
        return destination
    except Exception:
        partial.unlink(missing_ok=True)
        raise
