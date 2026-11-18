#!/usr/bin/env python3
"""Validate and inspect workspace-local presentation source profiles."""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from urllib.parse import parse_qsl, unquote, urlsplit

import yaml

DEFAULT_PATH = Path(".presentation-assembler/sources.yaml")
KINDS = {"file", "url", "connected_service", "user_input", "web", "other"}
AUTH_HEADER_KEY = "author" + "ization"
SECRET_PARAM_NAMES = {
    "auth", AUTH_HEADER_KEY, "token", "secret", "password", "private_key",
    "client_secret", "credential", "api_key", "apikey", "access_key",
    "session", "cookie", "oauth",
}
SECRET_KEY = re.compile(
    r"(auth|token|secret|password|private[_-]?key|credential|api[_-]?key|"
    r"access[_-]?key|session|author" r"ization|cookie|oauth)",
    re.I,
)
SECRET_VALUE = re.compile(
    r"(Bearer\s+\S+|Basic\s+\S+|ya29\.|AQAD-|gh[pousr]_[A-Za-z0-9_]{16,}|github_pat_[A-Za-z0-9_]{16,}|"
    r"(?:AKIA|ASIA)[A-Z0-9]{16}|xox[a-z]-[A-Za-z0-9-]{16,}|(?<![A-Za-z0-9])sk-[A-Za-z0-9_-]{16,}|-----BEGIN .*PRIVATE KEY-----|"
    r"https?://[^/\s:@]+:[^@\s/]+@|[?&](token|api[_-]?key|access[_-]?key|secret|session)=)", re.I)


def load(path: Path) -> dict:
    if not path.exists():
        raise ValueError(f"source profile not found: {path}")
    data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError("root must be a mapping")
    validate(data)
    return data


def nearest(start: Path | None = None) -> Path | None:
    current = (start or Path.cwd()).resolve()
    if current.is_file():
        current = current.parent
    for directory in (current, *current.parents):
        candidate = directory / DEFAULT_PATH
        if candidate.exists():
            return candidate
    return None


def validate(data: dict) -> None:
    allowed_root = {"version", "profiles"}
    allowed_profile = {"description", "use_when", "sources", "rules"}
    allowed_source = {"kind", "provider", "locator", "purpose"}
    allowed_rules = {"freshness", "owner"}

    def reject_unknown(mapping, allowed, where):
        unknown = set(mapping) - allowed
        if unknown:
            raise ValueError(f"{where}: unsupported keys: {', '.join(sorted(unknown))}")

    def walk(value, where="root"):
        if isinstance(value, dict):
            for key, child in value.items():
                if SECRET_KEY.search(str(key)):
                    raise ValueError(f"{where}: credential-like key is forbidden: {key}")
                walk(child, f"{where}.{key}")
        elif isinstance(value, list):
            for index, child in enumerate(value):
                walk(child, f"{where}[{index}]")
        else:
            decoded = unquote(str(value))
            if SECRET_VALUE.search(decoded):
                raise ValueError(f"{where}: credential-like value is forbidden")
            parsed = urlsplit(decoded)
            if parsed.username is not None or parsed.password is not None:
                raise ValueError(f"{where}: URL userinfo credentials are forbidden")
            pairs = parse_qsl(parsed.query, keep_blank_values=True)
            pairs += parse_qsl(parsed.fragment, keep_blank_values=True)
            for key, _ in pairs:
                normalized = unquote(key).lower().replace("-", "_")
                if normalized in SECRET_PARAM_NAMES:
                    raise ValueError(f"{where}: credential-like URL parameter is forbidden")
    walk(data)
    reject_unknown(data, allowed_root, "root")
    if data.get("version") != 1:
        raise ValueError("version must be 1")
    profiles = data.get("profiles")
    if not isinstance(profiles, dict) or not profiles:
        raise ValueError("profiles must be a non-empty mapping")
    for name, profile in profiles.items():
        if not re.fullmatch(r"[a-z0-9][a-z0-9_-]*", str(name)):
            raise ValueError(f"invalid profile name: {name}")
        if not isinstance(profile, dict) or not profile.get("description"):
            raise ValueError(f"profile {name}: description is required")
        reject_unknown(profile, allowed_profile, f"profile {name}")
        if "rules" in profile:
            if not isinstance(profile["rules"], dict):
                raise ValueError(f"profile {name}: rules must be a mapping")
            reject_unknown(profile["rules"], allowed_rules, f"profile {name}.rules")
        sources = profile.get("sources")
        if not isinstance(sources, list) or not sources:
            raise ValueError(f"profile {name}: sources must be non-empty")
        for index, source in enumerate(sources):
            label = f"profile {name}, source {index + 1}"
            if not isinstance(source, dict):
                raise ValueError(f"{label}: must be a mapping")
            reject_unknown(source, allowed_source, label)
            if source.get("kind") not in KINDS:
                raise ValueError(f"{label}: kind must be one of {sorted(KINDS)}")
            if not source.get("provider") or not source.get("locator"):
                raise ValueError(f"{label}: provider and locator are required")


def template() -> str:
    return """version: 1
profiles:
  team-weekly:
    description: Weekly product review
    use_when: Weekly metrics and decisions
    sources:
      - kind: file
        provider: csv
        locator: data/weekly.csv
        purpose: Verified product metrics
    rules:
      freshness: P7D
      owner: team
"""


def main() -> int:
    parser = argparse.ArgumentParser(description="Manage presentation source profiles")
    parser.add_argument("--path", type=Path)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("validate")
    sub.add_parser("list")
    show = sub.add_parser("show")
    show.add_argument("profile")
    match = sub.add_parser("match")
    match.add_argument("query")
    init = sub.add_parser("init")
    init.add_argument("--force", action="store_true")
    save = sub.add_parser("save")
    save.add_argument("--confirmed", action="store_true")
    save.add_argument("--name", required=True)
    save.add_argument("--description", required=True)
    save.add_argument("--kind", choices=sorted(KINDS), required=True)
    save.add_argument("--provider", required=True)
    save.add_argument("--locator", required=True)
    save.add_argument("--purpose", default="Presentation evidence")
    args = parser.parse_args()
    try:
        path = args.path or nearest() or DEFAULT_PATH
        if args.command == "init":
            if path.exists() and not args.force:
                raise ValueError(f"refusing to overwrite {path}; use --force")
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(template(), encoding="utf-8")
            print(path)
            return 0
        if args.command == "save":
            if not args.confirmed:
                raise ValueError("saving a source profile requires explicit --confirmed")
            data = load(path) if path.exists() else {"version": 1, "profiles": {}}
            if args.name in data["profiles"]:
                raise ValueError(f"profile already exists: {args.name}")
            data["profiles"][args.name] = {"description": args.description, "sources": [{
                "kind": args.kind, "provider": args.provider, "locator": args.locator,
                "purpose": args.purpose,
            }]}
            validate(data)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(yaml.safe_dump(data, allow_unicode=True, sort_keys=False), encoding="utf-8")
            print(path)
            return 0
        data = load(path)
        profiles = data["profiles"]
        if args.command == "validate":
            print(f"OK: {len(profiles)} profile(s)")
        elif args.command == "list":
            for name, profile in profiles.items():
                print(f"{name}\t{profile['description']}")
        elif args.command == "show":
            if args.profile not in profiles:
                raise ValueError(f"unknown profile: {args.profile}")
            print(yaml.safe_dump({args.profile: profiles[args.profile]}, allow_unicode=True, sort_keys=False))
        else:
            words = set(re.findall(r"[\w-]+", args.query.lower()))
            scores = {}
            for name, profile in profiles.items():
                haystack = " ".join((name, str(profile.get("description", "")), str(profile.get("use_when", "")))).lower()
                scores[name] = sum(1 for word in words if len(word) > 2 and word in haystack)
            best = max(scores.values(), default=0)
            winners = sorted(name for name, score in scores.items() if score == best and score > 0)
            if not winners:
                print("NO_MATCH")
                return 4
            if len(winners) > 1:
                print("AMBIGUOUS: " + ", ".join(winners))
                return 3
            print(winners[0])
        return 0
    except (OSError, ValueError, yaml.YAMLError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
