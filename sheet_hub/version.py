from __future__ import annotations

import re

import requests


APP_VERSION = "1.2.0"
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


def fetch_latest_release(timeout: int = 20) -> dict[str, str]:
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
    return {
        "tag": tag or f"v{version}",
        "version": version,
        "url": str(data.get("html_url") or RELEASES_URL),
        "name": str(data.get("name") or tag or version),
    }
