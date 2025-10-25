"""Networking client for the YBS Print Calander application."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Iterable, List, Mapping, Optional, Sequence

import requests
from bs4 import BeautifulSoup, Tag


class YBSError(Exception):
    """Base exception for YBS related errors."""


class AuthenticationError(YBSError):
    """Raised when authentication with the YBS portal fails."""


class NetworkError(YBSError):
    """Raised when the remote service cannot be reached."""


def _coerce_int(value: object) -> int | None:
    try:
        if isinstance(value, bool):
            return int(value)
        if isinstance(value, (int, float)):
            return int(value)
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _coerce_float(value: object) -> float | None:
    try:
        if isinstance(value, bool):
            return float(int(value))
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


@dataclass(eq=False)
class OrderRecord:
    """Represents a single order entry scraped from the manage page.

    The record now captures additional scheduling metadata used by the GUI when
    placing jobs onto the calendar.  ``color_count`` reflects the number of
    colors in the job, ``repeat_length`` stores the repeat length (in inches),
    ``press_time_minutes`` stores a computed or user-specified press time,
    ``press_model`` captures the configured press, and ``press_speed_ft_min``/
    ``press_speed_m_min`` store the calculated throughput values exported from
    the calculator dialog.
    """

    order_number: str
    company: str
    color_count: int | None = None
    repeat_length: float | None = None
    press_time_minutes: float | None = None
    press_model: str | None = None
    press_speed_ft_min: float | None = None
    press_speed_m_min: float | None = None

    def __post_init__(self) -> None:
        self.order_number = str(self.order_number or "").strip()
        self.company = str(self.company or "").strip()

        if self.color_count is not None:
            coerced = _coerce_int(self.color_count)
            self.color_count = coerced if coerced is not None else None

        if self.repeat_length is not None:
            coerced = _coerce_float(self.repeat_length)
            self.repeat_length = coerced if coerced is not None else None

        if self.press_time_minutes is not None:
            coerced = _coerce_float(self.press_time_minutes)
            self.press_time_minutes = coerced if coerced is not None else None

        if self.press_model is not None:
            text = str(self.press_model).strip()
            self.press_model = text or None

        if self.press_speed_ft_min is not None:
            coerced = _coerce_float(self.press_speed_ft_min)
            self.press_speed_ft_min = coerced if coerced is not None else None

        if self.press_speed_m_min is not None:
            coerced = _coerce_float(self.press_speed_m_min)
            self.press_speed_m_min = coerced if coerced is not None else None
        elif self.press_speed_ft_min is not None:
            self.press_speed_m_min = round(self.press_speed_ft_min / 3.28084, 4)

    def __eq__(self, other: object) -> bool:
        if not isinstance(other, OrderRecord):
            return NotImplemented
        return (
            self.order_number.lower(),
            self.company.lower(),
        ) == (
            other.order_number.lower(),
            other.company.lower(),
        )

    def copy(self) -> "OrderRecord":
        return OrderRecord(
            order_number=self.order_number,
            company=self.company,
            color_count=self.color_count,
            repeat_length=self.repeat_length,
            press_time_minutes=self.press_time_minutes,
            press_model=self.press_model,
            press_speed_ft_min=self.press_speed_ft_min,
            press_speed_m_min=self.press_speed_m_min,
        )

    def merge_metadata(self, other: "OrderRecord") -> "OrderRecord":
        """Return a copy that keeps identifiers but prefers ``other`` metadata."""

        return OrderRecord(
            order_number=self.order_number or other.order_number,
            company=self.company or other.company,
            color_count=other.color_count if other.color_count is not None else self.color_count,
            repeat_length=(
                other.repeat_length if other.repeat_length is not None else self.repeat_length
            ),
            press_time_minutes=(
                other.press_time_minutes
                if other.press_time_minutes is not None
                else self.press_time_minutes
            ),
            press_model=(
                other.press_model if other.press_model is not None else self.press_model
            ),
            press_speed_ft_min=(
                other.press_speed_ft_min
                if other.press_speed_ft_min is not None
                else self.press_speed_ft_min
            ),
            press_speed_m_min=(
                other.press_speed_m_min
                if other.press_speed_m_min is not None
                else self.press_speed_m_min
            ),
        )

    def metadata_summary(self) -> str:
        """Return a short textual summary of optional scheduling metadata."""

        parts: list[str] = []
        if self.press_model:
            parts.append(self.press_model)
        if self.color_count is not None:
            parts.append(f"{self.color_count}c")
        if self.repeat_length is not None:
            repeat = f"{self.repeat_length:g}" if self.repeat_length % 1 else f"{int(self.repeat_length)}"
            parts.append(f"{repeat}\" rpt")
        if self.press_time_minutes is not None:
            minutes = (
                f"{self.press_time_minutes:.1f}"
                if abs(self.press_time_minutes - round(self.press_time_minutes)) > 0.05
                else f"{int(round(self.press_time_minutes))}"
            )
            parts.append(f"{minutes} min")
        if self.press_speed_ft_min is not None:
            ft_speed = (
                f"{self.press_speed_ft_min:.1f}"
                if abs(self.press_speed_ft_min - round(self.press_speed_ft_min)) > 0.05
                else f"{int(round(self.press_speed_ft_min))}"
            )
            parts.append(f"{ft_speed} ft/min")
        if self.press_speed_m_min is not None:
            m_speed = (
                f"{self.press_speed_m_min:.1f}"
                if abs(self.press_speed_m_min - round(self.press_speed_m_min)) > 0.05
                else f"{int(round(self.press_speed_m_min))}"
            )
            parts.append(f"{m_speed} m/min")
        return ", ".join(parts)

    def label(self) -> str:
        """Return a human-friendly label combining the identifier and metadata."""

        order_number = self.order_number.strip()
        company = self.company.strip()
        if order_number and company:
            base = f"{order_number} - {company}"
        elif order_number:
            base = order_number
        elif company:
            base = company
        else:
            base = "Unnamed order"

        metadata = self.metadata_summary()
        return f"{base} [{metadata}]" if metadata else base

    def estimate_press_time(
        self,
        *,
        setup_minutes_per_color: float = 5.0,
        repeat_minutes_factor: float = 0.25,
    ) -> float | None:
        """Estimate the press time using the stored color and repeat information.

        The formula is intentionally simple and is designed to provide a starting
        point for manual adjustments: setup time scales by the number of colors
        and the repeat length contributes proportionally via ``repeat_minutes_factor``.
        The computed value is saved to ``press_time_minutes`` and returned.  ``None``
        is returned if no estimation can be made.
        """

        color_setup = 0.0
        repeat_component = 0.0

        if self.color_count is not None:
            color_setup = max(0.0, setup_minutes_per_color) * max(self.color_count, 0)
        if self.repeat_length is not None:
            repeat_component = max(0.0, repeat_minutes_factor) * max(self.repeat_length, 0.0)

        estimated = color_setup + repeat_component
        if estimated <= 0.0:
            self.press_time_minutes = None
            self.press_speed_ft_min = None
            self.press_speed_m_min = None
            return None

        self.press_time_minutes = round(estimated, 2)
        self.press_speed_ft_min = None
        self.press_speed_m_min = None
        return self.press_time_minutes

    def to_dict(self) -> dict[str, object]:
        data: dict[str, object] = {
            "order_number": self.order_number,
            "company": self.company,
        }
        if self.color_count is not None:
            data["color_count"] = self.color_count
        if self.repeat_length is not None:
            data["repeat_length"] = self.repeat_length
        if self.press_time_minutes is not None:
            data["press_time_minutes"] = self.press_time_minutes
        if self.press_model:
            data["press_model"] = self.press_model
        if self.press_speed_ft_min is not None:
            data["press_speed_ft_min"] = self.press_speed_ft_min
        if self.press_speed_m_min is not None:
            data["press_speed_m_min"] = self.press_speed_m_min
        return data

    @classmethod
    def from_dict(cls, payload: Mapping[str, object]) -> "OrderRecord":
        order_number = payload.get("order_number", "")
        company = payload.get("company", "")
        color_count = _coerce_int(payload.get("color_count"))
        repeat_length = _coerce_float(payload.get("repeat_length"))
        press_time = _coerce_float(payload.get("press_time_minutes"))
        press_model = payload.get("press_model")
        press_speed_ft = _coerce_float(
            payload.get("press_speed_ft_min")
            if "press_speed_ft_min" in payload
            else payload.get("press_speed_per_hour")
        )
        press_speed_m = _coerce_float(payload.get("press_speed_m_min"))
        return cls(
            order_number=str(order_number or ""),
            company=str(company or ""),
            color_count=color_count,
            repeat_length=repeat_length,
            press_time_minutes=press_time,
            press_model=press_model,
            press_speed_ft_min=press_speed_ft,
            press_speed_m_min=press_speed_m,
        )

    @classmethod
    def from_values(cls, values: object) -> "OrderRecord":
        if isinstance(values, OrderRecord):
            return values.copy()

        if isinstance(values, Mapping):
            return cls.from_dict(values)

        sequence: Sequence[object]
        if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
            sequence = values
        else:
            sequence = (values,) if values is not None else ()

        if len(sequence) == 1 and isinstance(sequence[0], OrderRecord):
            return sequence[0].copy()

        order_number = str(sequence[0]) if len(sequence) > 0 else ""
        company = str(sequence[1]) if len(sequence) > 1 else ""
        color_count = _coerce_int(sequence[2]) if len(sequence) > 2 else None
        repeat_length = _coerce_float(sequence[3]) if len(sequence) > 3 else None
        press_time = _coerce_float(sequence[4]) if len(sequence) > 4 else None
        raw_press_model = sequence[5] if len(sequence) > 5 else None
        press_speed_ft = _coerce_float(sequence[6]) if len(sequence) > 6 else None
        press_speed_m = _coerce_float(sequence[7]) if len(sequence) > 7 else None

        press_model: str | None
        if isinstance(raw_press_model, str):
            press_model = raw_press_model.strip() or None
        else:
            press_model = None
            if press_speed_ft is None:
                press_speed_ft = _coerce_float(raw_press_model)

        return cls(
            order_number=order_number,
            company=company,
            color_count=color_count,
            repeat_length=repeat_length,
            press_time_minutes=press_time,
            press_model=press_model,
            press_speed_ft_min=press_speed_ft,
            press_speed_m_min=press_speed_m,
        )


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
