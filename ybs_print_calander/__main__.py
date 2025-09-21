"""Entry point for running the GUI via ``python -m ybs_print_calander``."""

from .gui import launch_app


def main() -> None:
    launch_app()


if __name__ == "__main__":  # pragma: no cover - manual usage
    main()
