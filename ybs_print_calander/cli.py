"""Command line interface for the YBS Print Calander."""

from __future__ import annotations

import argparse
import os
import sys
from getpass import getpass
from typing import Iterable, Sequence

from .client import AuthenticationError, NetworkError, OrderRecord, YBSClient


def _prompt_for_missing(value: str | None, prompt: str, secret: bool = False) -> str:
    if value:
        return value
    if secret:
        return getpass(prompt)
    return input(prompt)


def _format_table(orders: Sequence[OrderRecord]) -> str:
    headers = ("Order#", "Company")
    column_widths = [len(header) for header in headers]

    for order in orders:
        column_widths[0] = max(column_widths[0], len(order.order_number))
        column_widths[1] = max(column_widths[1], len(order.company))

    divider = "+".join("-" * (width + 2) for width in column_widths)
    divider = f"+{divider}+"

    header_line = "|".join(
        f" {header.center(width)} " for header, width in zip(headers, column_widths)
    )
    header_line = f"|{header_line}|"

    lines = [divider, header_line, divider]

    for order in orders:
        row_line = "|".join(
            (
                f" {order.order_number.center(column_widths[0])} ",
                f" {order.company.center(column_widths[1])} ",
            )
        )
        lines.append(f"|{row_line}|")

    lines.append(divider)
    return "\n".join(lines)


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="YBS Print Calander CLI")
    parser.add_argument(
        "--username",
        "-u",
        help=(
            "Username used for the YBS portal. Can also be provided via the "
            "YBS_USERNAME environment variable"
        ),
    )
    parser.add_argument(
        "--password",
        "-p",
        help=(
            "Password used for the YBS portal. Can also be provided via the "
            "YBS_PASSWORD environment variable"
        ),
    )
    parser.add_argument(
        "--no-prompt",
        action="store_true",
        help="Fail immediately if credentials are not supplied via arguments",
    )

    args = parser.parse_args(list(argv) if argv is not None else None)

    username = args.username or os.environ.get("YBS_USERNAME")
    password = args.password or os.environ.get("YBS_PASSWORD")

    if args.no_prompt and (not username or not password):
        parser.error(
            "--no-prompt requires both --username and --password to be provided or "
            "for YBS_USERNAME/YBS_PASSWORD to be set"
        )

    if not username and not args.no_prompt:
        username = _prompt_for_missing(username, "Username: ")
    if not password and not args.no_prompt:
        password = _prompt_for_missing(password, "Password: ", secret=True)

    client = YBSClient()

    try:
        client.login(username, password)
    except AuthenticationError as exc:
        print(f"Login failed: {exc}", file=sys.stderr)
        return 1
    except NetworkError as exc:
        print(f"Network error: {exc}", file=sys.stderr)
        return 2

    try:
        orders = client.fetch_orders()
    except AuthenticationError as exc:
        print(f"Unable to retrieve orders: {exc}", file=sys.stderr)
        return 1
    except NetworkError as exc:
        print(f"Network error while fetching orders: {exc}", file=sys.stderr)
        return 2

    if not orders:
        print("No orders were returned from the manage page.")
        return 0

    print(_format_table(orders))
    return 0


if __name__ == "__main__":  # pragma: no cover - manual usage
    raise SystemExit(main())
