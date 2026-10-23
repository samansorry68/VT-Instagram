"""Optional Apify-backed public Instagram comments provider for V.T Instagram.

Set APIFY_API_TOKEN privately in the hosting environment to enable it.
No Instagram password or session cookie is required by this module.
"""
import json
import os
from datetime import datetime, timezone
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


class ApifyCommentsError(RuntimeError):
    pass


def fetch_comments(post_url: str, limit: int = 500) -> list[dict]:
    token = os.environ.get("APIFY_API_TOKEN", "").strip()
    if not token:
        raise ApifyCommentsError("APIFY_API_TOKEN is not configured")
    limit = min(max(int(limit), 1), 5000)
    endpoint = (
        "https://api.apify.com/v2/acts/apify~instagram-scraper/"
        "run-sync-get-dataset-items?token=" + token
        + "&timeout=120&memory=1024"
    )
    payload = json.dumps({
        "directUrls": [post_url],
        "resultsType": "comments",
        "resultsLimit": limit,
    }).encode("utf-8")
    request = Request(endpoint, data=payload, headers={
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": "VT-Instagram/1.0",
    }, method="POST")
    try:
        with urlopen(request, timeout=150) as response:
            items = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise ApifyCommentsError("Apify returned HTTP %d" % exc.code) from exc
    except (URLError, TimeoutError) as exc:
        raise ApifyCommentsError("Apify connection failed") from exc
    if not isinstance(items, list):
        raise ApifyCommentsError("Unexpected Apify response format")
    results = []
    for item in items[:limit]:
        if not isinstance(item, dict):
            continue
        owner = item.get("owner")
        username = item.get("ownerUsername") or item.get("username") or (
            owner.get("username", "") if isinstance(owner, dict) else ""
        )
        stamp = item.get("timestamp") or item.get("createdAt") or item.get("created_at") or ""
        if isinstance(stamp, (int, float)):
            stamp = datetime.fromtimestamp(stamp, timezone.utc).strftime("%Y-%m-%d %H:%M:%S")
        elif isinstance(stamp, str):
            stamp = stamp.replace("T", " ").replace("Z", "")[:19]
        results.append({
            "username": str(username or ""),
            "text": str(item.get("text") or item.get("comment") or item.get("commentText") or ""),
            "date": str(stamp or ""),
            "likes": item.get("likesCount", item.get("likes", 0)) or 0,
        })
    return results
