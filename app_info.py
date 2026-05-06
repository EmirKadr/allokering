"""Application identity and release metadata."""

APP_NAME = "Allokering"
APP_VERSION = "12.1.5"
APP_VERSION_DISPLAY = "12.1.5"
APP_BASE_TITLE = "Buffertpallar → Order-allokering (GUI)"
APP_TITLE = f"{APP_BASE_TITLE} — {APP_VERSION_DISPLAY}"
GITHUB_REPO = "EmirKadr/allokering"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases"
UPDATE_DISABLED_ENV = "ALLOKERING_DISABLE_UPDATE_CHECK"

# Owner-only analytics config.
# By default events are stored locally in %APPDATA%\allokering\analytics and
# read by analytics_dashboard.py. If you later want to aggregate multiple
# users without a server, set ANALYTICS_LOCAL_STORAGE_DIR to a shared folder
# before distributing the app.
ANALYTICS_ENABLED_DEFAULT = True
ANALYTICS_LOCAL_STORAGE_DIR = ""
ANALYTICS_STORAGE_DIR_ENV = "ALLOKERING_ANALYTICS_STORAGE_DIR"
ANALYTICS_POSTHOG_HOST = "https://us.i.posthog.com"
ANALYTICS_POSTHOG_PROJECT_API_KEY = ""
ANALYTICS_ENABLED_ENV = "ALLOKERING_ANALYTICS_ENABLED"
ANALYTICS_HOST_ENV = "ALLOKERING_ANALYTICS_HOST"
ANALYTICS_PROJECT_API_KEY_ENV = "ALLOKERING_ANALYTICS_PROJECT_API_KEY"
