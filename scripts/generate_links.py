"""
generate_links.py
-----------------
Queries Microsoft Graph API to find each lesson PDF by filename
and updates links.json with fresh webUrls.

Required environment variables:
  TENANT_ID      - Azure AD tenant ID
  CLIENT_ID      - App registration client ID
  CLIENT_SECRET  - App registration client secret
  DRIVE_ID       - SharePoint drive ID (optional fallback)
"""

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

TENANT_ID = os.environ["TENANT_ID"]
CLIENT_ID = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
DRIVE_ID = os.environ.get("DRIVE_ID", "")  # optional, improves search speed

GRAPH_URL = "https://graph.microsoft.com/v1.0"

LESSONS = [
    {
        "lesson": 1,
        "title": "Lesson 1",
        "desc": "Integration Fundamentals and APIs",
        "filename": "Lesson 1 - Integration Fundamentals and APIs.pdf",
    },
    {
        "lesson": 2,
        "title": "Lesson 2",
        "desc": "OData and the SAP World",
        "filename": "Lesson 2 - OData and the SAP World.pdf",
    },
    {
        "lesson": 3,
        "title": "Lesson 3",
        "desc": "Introduction to SAP BTP",
        "filename": "Lesson 3 - Introduction to SAP BTP.pdf",
    },
    {
        "lesson": 4,
        "title": "Lesson 4",
        "desc": "First iFlow with an Open API",
        "filename": "Lesson 4 - First iFlow with an Open API.pdf",
    },
    {
        "lesson": 5,
        "title": "Lesson 5",
        "desc": "Connecting to SAP SuccessFactors",
        "filename": "Lesson 5 - Connecting to SAP SuccessFactors.pdf",
    },
    {
        "lesson": 6,
        "title": "Lesson 6",
        "desc": "Best Practices and Final Project",
        "filename": "Lesson 6 - Best Practices and Final Project.pdf",
    },
]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def http_get(url: str, token: str) -> dict:
    req = urllib.request.Request(url)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Accept", "application/json")
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as exc:
        body = exc.read().decode(errors="replace")
        print(f"  HTTP {exc.code} on GET {url[:100]}:\n  {body[:300]}", file=sys.stderr)
        raise


def http_post(url: str, token: str, payload: dict) -> dict:
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json")
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as exc:
        body = exc.read().decode(errors="replace")
        print(f"  HTTP {exc.code} on POST {url[:100]}:\n  {body[:300]}", file=sys.stderr)
        raise


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


def get_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    params = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    data = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    try:
        with urllib.request.urlopen(req) as resp:
            result = json.loads(resp.read())
    except urllib.error.HTTPError as exc:
        body = exc.read().decode(errors="replace")
        print(f"Auth failed ({exc.code}): {body[:500]}", file=sys.stderr)
        sys.exit(1)

    token = result.get("access_token")
    if not token:
        print(f"No access_token in response: {result}", file=sys.stderr)
        sys.exit(1)
    return token


# ---------------------------------------------------------------------------
# Search strategies
# ---------------------------------------------------------------------------


def search_in_drive(token: str, filename: str) -> dict | None:
    """Search within the configured drive (fast, scoped)."""
    if not DRIVE_ID:
        return None
    encoded = urllib.parse.quote(filename)
    url = (
        f"{GRAPH_URL}/drives/{DRIVE_ID}/root/search(q='{encoded}')"
        "?$select=id,name,webUrl&$top=25"
    )
    try:
        result = http_get(url, token)
    except urllib.error.HTTPError:
        return None

    for item in result.get("value", []):
        if item.get("name", "").lower() == filename.lower():
            return item

    # Handle pagination
    next_link = result.get("@odata.nextLink")
    while next_link:
        try:
            result = http_get(next_link, token)
        except urllib.error.HTTPError:
            break
        for item in result.get("value", []):
            if item.get("name", "").lower() == filename.lower():
                return item
        next_link = result.get("@odata.nextLink")

    return None


def search_global(token: str, filename: str) -> dict | None:
    """Search across all accessible SharePoint content (broader)."""
    url = f"{GRAPH_URL}/search/query"
    payload = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": f'"{filename}"'},
                "fields": ["name", "webUrl", "id"],
                "size": 10,
            }
        ]
    }
    try:
        result = http_post(url, token, payload)
    except urllib.error.HTTPError:
        return None

    containers = result.get("value", [{}])[0].get("hitsContainers", [])
    for container in containers:
        for hit in container.get("hits", []):
            resource = hit.get("resource", {})
            if resource.get("name", "").lower() == filename.lower():
                return {
                    "id": resource.get("id"),
                    "name": resource.get("name"),
                    "webUrl": resource.get("webUrl"),
                }
    return None


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def load_existing() -> dict:
    """Return {lesson_number: url} from current links.json as fallback."""
    try:
        with open("links.json", encoding="utf-8") as f:
            return {item["lesson"]: item.get("url", "") for item in json.load(f)}
    except (FileNotFoundError, json.JSONDecodeError, KeyError):
        return {}


def main() -> None:
    print("Obtaining access token…")
    token = get_token()
    print("Token obtained.\n")

    existing = load_existing()
    output = []
    not_found = []

    for lesson in LESSONS:
        fn = lesson["filename"]
        print(f"[{lesson['lesson']}/6] Searching: {fn}")

        item = None

        # 1. Try scoped drive search
        item = search_in_drive(token, fn)
        if item:
            print(f"  ✓ Found in drive: {item.get('webUrl', '')[:90]}")
        else:
            # 2. Fallback to tenant-wide search
            print("  → Not found in drive, trying global search…")
            item = search_global(token, fn)
            if item:
                print(f"  ✓ Found globally: {item.get('webUrl', '')[:90]}")

        if item and item.get("webUrl"):
            url = item["webUrl"]
        else:
            url = existing.get(lesson["lesson"], "")
            if url:
                print(f"  ⚠ Not found — keeping existing URL from links.json")
            else:
                print(f"  ✗ Not found and no fallback URL available", file=sys.stderr)
                not_found.append(fn)

        output.append(
            {
                "lesson": lesson["lesson"],
                "title": lesson["title"],
                "desc": lesson["desc"],
                "url": url,
            }
        )

    with open("links.json", "w", encoding="utf-8") as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
        f.write("\n")

    print("\nlinks.json written.")

    if not_found:
        print(
            f"\n⚠ Warning: {len(not_found)} file(s) not found in SharePoint:",
            file=sys.stderr,
        )
        for fn in not_found:
            print(f"  - {fn}", file=sys.stderr)
        print(
            "Ensure the filenames match exactly and the service principal has "
            "Files.Read.All / Sites.Read.All permissions.",
            file=sys.stderr,
        )
        sys.exit(1)  # fail the workflow so the maintainer is notified


if __name__ == "__main__":
    main()
