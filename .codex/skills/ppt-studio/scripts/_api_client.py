#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, Optional


DEFAULT_API_BASE = os.environ.get("PPT_API_BASE", "http://10.1.3.127:6414").rstrip("/")


def get_json(path: str, query: Dict[str, Any], api_base: Optional[str] = None) -> Dict[str, Any]:
    base = (api_base or DEFAULT_API_BASE).rstrip("/")
    params = "&".join(f"{key}={urllib.parse.quote(str(value))}" for key, value in query.items())
    url = f"{base}{path}?{params}"
    request = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=3600) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{url} failed: {text}") from exc


def post_json(path: str, payload: Dict[str, Any], api_base: Optional[str] = None) -> Dict[str, Any]:
    base = (api_base or DEFAULT_API_BASE).rstrip("/")
    url = f"{base}{path}"
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=3600) as response:
            raw = response.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except urllib.error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{url} failed: {text}") from exc

