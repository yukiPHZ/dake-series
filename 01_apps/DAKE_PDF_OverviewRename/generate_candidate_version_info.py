"""Generate Windows file version for this source build, not the published Store view."""

from pathlib import Path
import sys

APP_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(APP_DIR.parents[1] / "tools"))

from generate_version_info import build_version_info, read_meta  # noqa: E402
from main import APP_VERSION  # noqa: E402


def main() -> None:
    meta = read_meta(APP_DIR)
    meta["version"] = APP_VERSION
    (APP_DIR / "version_info.txt").write_text(build_version_info(meta), encoding="utf-8")


if __name__ == "__main__":
    main()
