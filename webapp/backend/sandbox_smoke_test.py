from __future__ import annotations

try:
    from .test_regression_flow import main
except ImportError:
    from test_regression_flow import main


if __name__ == "__main__":
    main()
