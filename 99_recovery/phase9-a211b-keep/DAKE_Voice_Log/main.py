# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse

from app.ui.main_window import VoiceLogApp


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--launch-check", action="store_true")
    parser.add_argument("--smoke-seconds", type=float, default=None)
    args = parser.parse_args()
    smoke_seconds = args.smoke_seconds
    if args.launch_check and smoke_seconds is None:
        smoke_seconds = 3.0
    app = VoiceLogApp(start_services=not args.launch_check, smoke_seconds=smoke_seconds)
    app.mainloop()


if __name__ == "__main__":
    main()

