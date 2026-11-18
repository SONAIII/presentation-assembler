#!/usr/bin/env python3
"""Security gate for custom HTML before it reaches Chromium.

Custom slides are declarative layout documents, not web applications.  Reject
active content and network-capable references so user-provided HTML cannot run
scripts or make outbound requests during measurement and preview.
"""
from __future__ import annotations

from html.parser import HTMLParser
import re
from pathlib import Path
from urllib.parse import urlsplit


FORBIDDEN_TAGS = {
    "script", "iframe", "object", "embed", "link", "base", "form",
    "input", "button", "textarea", "select", "video", "audio", "source",
    "track", "foreignobject",
}
URL_ATTRIBUTES = {"src", "href", "xlink:href", "action", "formaction", "poster", "data"}
SVG_URL_ATTRIBUTES = {"filter", "mask", "clip-path", "fill", "stroke", "cursor"}
REMOTE_SCHEMES = {"http", "https", "ftp", "ws", "wss", "javascript"}
CSS_IMPORT_RE = re.compile(r"@import\b", re.IGNORECASE)
CSS_URL_RE = re.compile(r"url\(\s*(['\"]?)(.*?)\1\s*\)", re.IGNORECASE | re.DOTALL)


class UnsafeHtmlError(ValueError):
    pass


def _validate_reference(value: str, context: str) -> None:
    value = value.strip()
    if not value or value.startswith("#"):
        return
    parsed = urlsplit(value)
    scheme = parsed.scheme.lower()
    if scheme in REMOTE_SCHEMES or value.startswith("//"):
        raise UnsafeHtmlError(f"{context}: external/active URL is forbidden: {value[:120]}")
    if scheme == "data" and not value.lower().startswith("data:image/"):
        raise UnsafeHtmlError(f"{context}: only data:image URLs are allowed")
    if scheme not in {"", "file", "data"}:
        raise UnsafeHtmlError(f"{context}: unsupported URL scheme: {scheme}")


def _validate_css(css: str, context: str) -> None:
    if "\\" in css:
        raise UnsafeHtmlError(f"{context}: CSS escapes are forbidden")
    if CSS_IMPORT_RE.search(css):
        raise UnsafeHtmlError(f"{context}: CSS @import is forbidden")
    for match in CSS_URL_RE.finditer(css):
        _validate_reference(match.group(2), f"{context} url()")


class _GateParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_style = False

    def handle_starttag(self, tag: str, attrs) -> None:
        lower = tag.lower()
        if lower in FORBIDDEN_TAGS:
            raise UnsafeHtmlError(f"forbidden HTML tag: <{tag}>")
        if lower == "meta":
            attr_map = {str(k).lower(): str(v or "") for k, v in attrs}
            if attr_map.get("http-equiv", "").lower() == "refresh":
                raise UnsafeHtmlError("meta refresh is forbidden")
        self.in_style = lower == "style"
        for name, raw_value in attrs:
            attr = str(name).lower()
            value = str(raw_value or "")
            if attr.startswith("on"):
                raise UnsafeHtmlError(f"event handler is forbidden: {name}")
            if attr in URL_ATTRIBUTES:
                _validate_reference(value, f"attribute {name}")
            if attr == "srcset":
                for candidate in value.split(","):
                    _validate_reference(candidate.strip().split()[0], f"attribute {name}")
            if attr in SVG_URL_ATTRIBUTES:
                _validate_css(value, f"attribute {name}")
            if attr == "style":
                _validate_css(value, "style attribute")

    def handle_startendtag(self, tag: str, attrs) -> None:
        self.handle_starttag(tag, attrs)
        self.in_style = False

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "style":
            self.in_style = False

    def handle_data(self, data: str) -> None:
        if self.in_style:
            _validate_css(data, "style block")


def validate_html_text(html: str) -> None:
    parser = _GateParser()
    try:
        parser.feed(html)
        parser.close()
    except UnsafeHtmlError:
        raise
    except Exception as exc:
        raise UnsafeHtmlError(f"cannot safely parse custom HTML: {exc}") from exc


def validate_html_file(path: str | Path) -> str:
    html = Path(path).read_text(encoding="utf-8")
    validate_html_text(html)
    return html


CSP = (
    "default-src 'none'; img-src file: data:; style-src 'unsafe-inline'; "
    "font-src file: data:; script-src 'none'; connect-src 'none'; "
    "media-src 'none'; object-src 'none'; frame-src 'none'; base-uri 'none'; "
    "form-action 'none'"
)


def inject_csp(html: str) -> str:
    meta = f'<meta http-equiv="Content-Security-Policy" content="{CSP}">'
    if re.search(r"<head(?:\s[^>]*)?>", html, flags=re.IGNORECASE):
        return re.sub(r"(<head(?:\s[^>]*)?>)", r"\1" + meta, html,
                      count=1, flags=re.IGNORECASE)
    return meta + html
