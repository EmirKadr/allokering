from __future__ import annotations

from update_service import (
    UpdateInfo,
    check_for_update,
    download_update_installer,
    is_newer_version,
)


class FakeResponse:
    def __init__(self, data=None, status_code=200, content=b"", headers=None):
        self._data = data or {}
        self.status_code = status_code
        self._content = content
        self.headers = headers or {}

    def raise_for_status(self):
        if self.status_code >= 400:
            raise RuntimeError(f"HTTP {self.status_code}")

    def json(self):
        return self._data

    def iter_content(self, chunk_size=1):
        for start in range(0, len(self._content), chunk_size):
            yield self._content[start:start + chunk_size]


class FakeSession:
    def __init__(self, response):
        self.response = response
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        return self.response


def test_is_newer_version_handles_v_prefix():
    assert is_newer_version("v12.2.0", "12.1.9") is True


def test_is_newer_version_pads_missing_parts():
    assert is_newer_version("1.0", "1.0.0") is False


def test_check_for_update_returns_none_for_current_version():
    response = FakeResponse({"tag_name": "v12.1.3", "assets": []})
    assert check_for_update(current_version="12.1.3", session=FakeSession(response)) is None


def test_check_for_update_finds_setup_asset():
    response = FakeResponse(
        {
            "tag_name": "v12.2.0",
            "html_url": "https://example.test/releases/v12.2.0",
            "assets": [
                {"name": "Allokering-12.2.0-win64.zip", "browser_download_url": "zip"},
                {"name": "Allokering-12.2.0-Setup.exe", "browser_download_url": "exe"},
            ],
        }
    )

    info = check_for_update(current_version="12.1.3", session=FakeSession(response))

    assert info is not None
    assert info.version == "12.2.0"
    assert info.installer_name == "Allokering-12.2.0-Setup.exe"
    assert info.installer_url == "exe"


def test_download_update_installer_writes_file(tmp_path):
    response = FakeResponse(content=b"installer", headers={"content-length": "9"})
    info = UpdateInfo(
        version="12.2.0",
        tag_name="v12.2.0",
        release_url="https://example.test/release",
        installer_url="https://example.test/setup.exe",
        installer_name="Allokering-12.2.0-Setup.exe",
    )
    progress = []

    path = download_update_installer(
        info,
        target_dir=tmp_path,
        session=FakeSession(response),
        progress_cb=progress.append,
    )

    assert path.name == "Allokering-12.2.0-Setup.exe"
    assert path.read_bytes() == b"installer"
    assert progress[-1] == 100
