import re
from datetime import datetime
from zoneinfo import ZoneInfo
from urllib.parse import urlparse

from .settings import DATE_TZ, DATE_FMT, NOISE_SUBSTRINGS

def uk_today_str() -> str:
    return datetime.now(ZoneInfo(DATE_TZ)).strftime(DATE_FMT)

def clean_slug_to_name(slug: str) -> str:
    return slug.replace("-", " ").strip().title()

def fallback_page_name_from_url(url: str) -> str:
    path = urlparse(url).path.strip("/")
    parts = [p for p in path.split("/") if p]
    try:
        i = parts.index("destinations")
        if len(parts) > i + 2:
            return clean_slug_to_name(parts[i + 2])
    except ValueError:
        pass
    return clean_slug_to_name(parts[-1] if parts else (urlparse(url).hostname or "Page"))

def normalise_keep_newlines(s: str) -> str:
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"[ \t]*\n[ \t]*", "\n", s)
    return s

def is_noise(text: str) -> bool:
    t = (text or "").strip().lower()
    if not t:
        return False
    return any(sub in t for sub in NOISE_SUBSTRINGS)

def safe_filename(name: str, maxlen: int = 120) -> str:
    # collapse any whitespace/newlines to single spaces
    name = re.sub(r"\s+", " ", name)
    # remove characters that break downloads
    name = re.sub(r'[\\/*?:"<>|]+', "", name)
    # commas are legal but can confuse some agents – make safer
    name = name.replace(",", "")
    # trim length and trailing dots/spaces
    return (name[:maxlen]).rstrip(". ")
