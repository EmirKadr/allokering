"""GitHub Release based update checks and installer downloads."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import tempfile
from typing import Callable, Optional

import requests

from app_info import APP_NAME, APP_VERSION, GITHUB_REPO


ProgressCallback = Callable[[int], None]
StopFlag = Callable[[], bool]


@dataclass(frozen=True)
class UpdateInfo:
    version: str
    tag_name: str
    release_url: str
    installer_url: str
    installer_name: str
    notes: str = ""


class UpdateError(RuntimeError):
    """Raised when update metadata or installer download fails."""


def _version_parts(version: str) -> tuple[int, ...]:
    cleaned = version.strip().lstrip("vV").split("+", 1)[0].split("-", 1)[0]
    numbers = [int(part) for part in re.findall(r"\d+", cleaned)]
    return tuple(numbers or [0])


def is_newer_version(latest: str, current: str) -> bool:
    latest_parts = _version_parts(latest)
    current_parts = _version_parts(current)
    width = max(len(latest_parts), len(current_parts), 3)
    latest_parts += (0,) * (width - len(latest_parts))
    current_parts += (0,) * (width - len(current_parts))
    return latest_parts > current_parts


def _find_installer_asset(assets: list[dict]) -> dict:
    app_name = APP_NAME.lower()
    setup_assets = [
        asset
        for asset in assets
        if str(asset.get("name", "")).lower().endswith(".exe")
        and "setup" in str(asset.get("name", "")).lower()
    ]
    preferred = [
        asset
        for asset in setup_assets
        if app_name in str(asset.get("name", "")).lower()
    ]
    return (preferred or setup_assets or [{}])[0]


def check_for_update(
    current_version: str = APP_VERSION,
    repo: str = GITHUB_REPO,
    timeout: int = 8,
    session=None,
) -> Optional[UpdateInfo]:
    """Return update info when GitHub has a newer release, otherwise None."""
    http = session or requests
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    try:
        response = http.get(
            url,
            headers={"Accept": "application/vnd.github+json"},
            timeout=timeout,
        )
    except requests.RequestException as exc:
        raise UpdateError(str(exc)) from exc

    if getattr(response, "status_code", None) == 404:
        return None
    try:
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, ValueError) as exc:
        raise UpdateError(str(exc)) from exc

    tag_name = str(data.get("tag_name") or "")
    latest_version = tag_name.lstrip("vV")
    if not latest_version or not is_newer_version(latest_version, current_version):
        return None

    asset = _find_installer_asset(data.get("assets") or [])
    return UpdateInfo(
        version=latest_version,
        tag_name=tag_name,
        release_url=str(
            data.get("html_url") or f"https://github.com/{repo}/releases/latest"
        ),
        installer_url=str(asset.get("browser_download_url") or ""),
        installer_name=str(asset.get("name") or f"{APP_NAME}-{latest_version}-Setup.exe"),
        notes=str(data.get("body") or ""),
    )


def download_update_installer(
    info: UpdateInfo,
    target_dir: Optional[Path] = None,
    timeout: int = 60,
    session=None,
    progress_cb: Optional[ProgressCallback] = None,
    stop_flag: Optional[StopFlag] = None,
) -> Path:
    """Download the installer asset and return the local file path."""
    if not info.installer_url:
        raise UpdateError("Release saknar Setup.exe-asset.")

    root = Path(target_dir) if target_dir else Path(tempfile.gettempdir())
    root.mkdir(parents=True, exist_ok=True)
    target = root / info.installer_name

    http = session or requests
    try:
        response = http.get(info.installer_url, stream=True, timeout=timeout)
        response.raise_for_status()
    except requests.RequestException as exc:
        raise UpdateError(str(exc)) from exc

    total = int(response.headers.get("content-length") or 0)
    downloaded = 0
    with open(target, "wb") as fh:
        for chunk in response.iter_content(chunk_size=1024 * 128):
            if stop_flag and stop_flag():
                raise UpdateError("Nedladdningen avbröts.")
            if not chunk:
                continue
            fh.write(chunk)
            downloaded += len(chunk)
            if progress_cb and total:
                progress_cb(min(100, int(downloaded * 100 / total)))

    if progress_cb:
        progress_cb(100)
    return target
