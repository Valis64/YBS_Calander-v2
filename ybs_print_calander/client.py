"""Networking client for the YBS Print Calander application."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable, List, Optional

import requests
from bs4 import BeautifulSoup, Tag


class YBSError(Exception):
    """Base exception for YBS related errors."""


class AuthenticationError(YBSError):
    """Raised when authentication with the YBS portal fails."""


class NetworkError(YBSError):
    """Raised when the remote service cannot be reached."""


@dataclass
class OrderRecord:
    """Represents a single order entry scraped from the manage page."""

    order_number: str
    company: str


class YBSClient:
    """Simple HTTP client responsible for logging in and scraping orders."""

    LOGIN_URL = "https://www.ybsnow.com/index.php"
    MANAGE_URL = "https://www.ybsnow.com/manage.html"

    def __init__(self, session: Optional[requests.Session] = None) -> None:
        self.session = session or requests.Session()
        self.session.headers.setdefault(
            "User-Agent",
            "YBS Print Calander/1.0 (+https://www.ybsnow.com/)",
        )

    def login(self, username: str, password: str) -> bool:
        """Attempt to authenticate with the YBS website.

        Args:
            username: The username or email address used to sign in.
            password: The account password.

        Returns:
            ``True`` if the login appears to be successful.

        Raises:
            AuthenticationError: If the credentials are rejected by the server.
            NetworkError: If the network request cannot be completed.
        """

        payload = {
            "email": username,
            "password": password,
            "action": "signin",
        }

        try:
            response = self.session.post(self.LOGIN_URL, data=payload, timeout=10)
            response.raise_for_status()
        except requests.RequestException as exc:  # pragma: no cover - defensive
            raise NetworkError("Failed to reach the YBS login page.") from exc

        # Verify authentication by attempting to load the manage page.
        try:
            manage_response = self.session.get(self.MANAGE_URL, timeout=10)
            manage_response.raise_for_status()
        except requests.RequestException as exc:  # pragma: no cover - defensive
            raise NetworkError("Failed to verify login with the manage page.") from exc

        if self._is_login_page(manage_response.text):
            raise AuthenticationError("Login failed. Please verify your username and password.")

        return True

    def fetch_orders(self) -> List[OrderRecord]:
        """Fetch and parse the orders from the manage page."""

        try:
            response = self.session.get(self.MANAGE_URL, timeout=10)
            response.raise_for_status()
        except requests.RequestException as exc:  # pragma: no cover - defensive
            raise NetworkError("Failed to retrieve the orders page.") from exc

        if self._is_login_page(response.text):
            raise AuthenticationError("Cannot fetch orders without logging in first.")

        return list(self._parse_orders(response.text))

    def _is_login_page(self, html: str) -> bool:
        lowered = html.lower()
        return "id=\"signin\"" in lowered or "name=\"signin\"" in lowered

    def _parse_orders(self, html: str) -> Iterable[OrderRecord]:
        soup = BeautifulSoup(html, "html.parser")

        for row in soup.find_all("tr"):
            move_cell = row.find("td", class_="move")
            details_cell = row.find("td", class_=re.compile(r"\bdetails\b"))
            if move_cell is None or details_cell is None:
                continue

            order_number = self._extract_order_number(move_cell.get_text(" ", strip=True))
            company = self._extract_company(details_cell)

            if order_number and company:
                yield OrderRecord(order_number=order_number, company=company)

    def _extract_order_number(self, text: str) -> Optional[str]:
        match = re.search(r"\b(\d+)\b", text)
        if match:
            return match.group(1)
        return None

    def _extract_company(self, cell: Tag) -> Optional[str]:
        first_paragraph = cell.find("p")
        if first_paragraph:
            return first_paragraph.get_text(strip=True)
        return cell.get_text(strip=True) or None
