"""Windows EXE launcher for the Streamlit patent tool."""

from __future__ import annotations

import os
import sys

from streamlit.web import cli as stcli


def main() -> None:
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")
    sys.argv = ["streamlit", "run", app_path, "--server.headless=true"]
    raise SystemExit(stcli.main())


if __name__ == "__main__":
    main()
