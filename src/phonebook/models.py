"""Core data models and normalization helpers."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field


NBSP = "\xa0"


@dataclass(frozen=True)
class HubRecord:
    id: int
    name: str
    slug: str


@dataclass
class ContactRecord:
    lastname: str = ""
    firstname: str = ""
    relation: str = ""
    phone: str = ""
    mobile: str = ""


@dataclass
class CustomerRecord:
    lastname: str = ""
    firstname: str = ""
    carelevel: str = ""
    phone: str = ""
    mobile: str = ""
    external_id: str = ""
    contacts: list[ContactRecord] = field(default_factory=list)


def clean(value: object) -> str:
    if value is None:
        return ""
    text = str(value).replace(NBSP, " ").strip()
    return re.sub(r"\s+", " ", text)


def normalize_phone(value: object) -> str:
    text = clean(value)
    if not text:
        return ""
    text = text.replace("(0)", "0")
    return re.sub(r"[^0-9+/ -]", "", text).strip()


def normalized_phone_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def split_name(value: str) -> tuple[str, str]:
    value = clean(value)
    if not value:
        return "", ""
    if "," in value:
        last, first = value.split(",", 1)
        return clean(last), clean(first)
    parts = value.split()
    if len(parts) == 1:
        return parts[0], ""
    return " ".join(parts[:-1]), parts[-1]


def compute_customer_key(rec: CustomerRecord) -> str:
    raw = "|".join(
        [
            rec.external_id.lower(),
            rec.lastname.lower(),
            rec.firstname.lower(),
            normalized_phone_digits(rec.phone),
            normalized_phone_digits(rec.mobile),
        ]
    )
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()
