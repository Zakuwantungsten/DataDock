try:
    from .gui import run_app
except ImportError:  # pragma: no cover - supports running as a script
    from gui import run_app


if __name__ == "__main__":
    run_app()
