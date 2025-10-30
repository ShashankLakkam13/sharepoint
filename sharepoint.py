#!/usr/bin/env python3
"""SharePoint ASPX crawler.

This module implements a SharePoint crawler that enumerates drives and lists
in a SharePoint Online site, looking for `.aspx` pages.  It supports both
device-code (delegated) and application (client credential) authentication via
Microsoft Graph and optionally downloads discovered pages to the local
filesystem.

The original internal tool that inspired this module stored Azure client
credentials directly in the source file.  To keep this repository safe we load
configuration from environment variables instead.  See `load_config()` below
for the complete list of variables that can be provided.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import pathlib
import sys
import time
from typing import Dict, Iterable, List, Optional
from urllib.parse import urlsplit

import importlib.util

import requests
import urllib3
from msal import ConfidentialClientApplication, PublicClientApplication


# ---------------------------------------------------------------------------
# Optional corporate certificate support
# ---------------------------------------------------------------------------

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

if importlib.util.find_spec("pip_system_certs.wrapt_requests") is not None:
    import pip_system_certs.wrapt_requests  # type: ignore  # noqa: F401

    print("🔧 Corporate certificate trust initialized...")


# ---------------------------------------------------------------------------
# Constants and configuration helpers
# ---------------------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE_APP = ["https://graph.microsoft.com/.default"]
GRAPH_SCOPE_USER = ["Sites.Read.All", "Files.Read.All", "User.Read"]
RETRY_STATUS = {429, 500, 502, 503, 504}

ENV_PREFIX = "SHAREPOINT_"

FILE_KEY_ALIASES = {
    "TENANTID": "AZURE_TENANT_ID",
    "AZURETENANTID": "AZURE_TENANT_ID",
    "CLIENTID": "AZURE_CLIENT_ID",
    "AZURECLIENTID": "AZURE_CLIENT_ID",
    "CLIENTSECRET": "AZURE_CLIENT_SECRET",
    "AZURECLIENTSECRET": "AZURE_CLIENT_SECRET",
    "SPO_HOSTNAME": "SPO_HOSTNAME",
    "SPOHOSTNAME": "SPO_HOSTNAME",
    "SPO_SITE_PATH": "SPO_SITE_PATH",
    "SPOSITEPATH": "SPO_SITE_PATH",
    "SPO_SITE_URL": "SPO_SITE_URL",
    "SPOSITEURL": "SPO_SITE_URL",
    "SITEURL": "SPO_SITE_URL",
}


def bool_from_env(value, *, default: bool) -> bool:
    """Convert an environment variable string to ``bool``."""

    if value is None:
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def load_config_file(path: pathlib.Path) -> Dict[str, object]:
    """Load configuration values from a JSON file."""

    if not path.exists():
        raise ValueError(f"Configuration file not found: {path}")

    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
    except json.JSONDecodeError as exc:  # pragma: no cover - user input
        raise ValueError(f"Invalid JSON in configuration file: {exc}") from exc

    if not isinstance(data, dict):
        raise ValueError("Configuration file must contain a JSON object at the top level.")

    normalized: Dict[str, object] = {}
    for key, value in data.items():
        upper_key = str(key).upper()
        mapped_key = FILE_KEY_ALIASES.get(upper_key, upper_key)
        normalized[mapped_key] = value

    return normalized


def load_config(*, config_path: Optional[str] = None) -> Dict[str, object]:
    """Load configuration from the environment and optional JSON file."""

    file_values: Dict[str, object] = {}

    if config_path is None:
        config_path = os.getenv(f"{ENV_PREFIX}CONFIG_FILE")

    if config_path:
        file_values = load_config_file(pathlib.Path(config_path))

    def env(name: str, default: Optional[str] = None):
        if f"{ENV_PREFIX}{name}" in os.environ:
            return os.getenv(f"{ENV_PREFIX}{name}", default)
        value = file_values.get(name, default)
        if value is None:
            return default
        return value

    auth_mode = env("AUTH_MODE", "device").lower()
    if auth_mode not in {"device", "application"}:
        raise ValueError(
            "AUTH_MODE must be either 'device' or 'application' (case insensitive)."
        )

    site_url_value = env("SPO_SITE_URL")

    config: Dict[str, object] = {
        "AUTH_MODE": auth_mode,
        "AZURE_TENANT_ID": str(env("AZURE_TENANT_ID", "") or "").strip(),
        "AZURE_CLIENT_ID": str(env("AZURE_CLIENT_ID", "") or "").strip(),
        "AZURE_CLIENT_SECRET": str(env("AZURE_CLIENT_SECRET", "") or ""),
        "SPO_HOSTNAME": str(env("SPO_HOSTNAME", "") or "").strip(),
        "SPO_SITE_PATH": str(env("SPO_SITE_PATH", "") or "").strip() or "/",
        "DOWNLOAD_ASPX": bool_from_env(env("DOWNLOAD_ASPX"), default=True),
        "DOWNLOAD_DIR": pathlib.Path(env("DOWNLOAD_DIR", "./aspx_dump")),
        "EXPORT_CSV": pathlib.Path(env("EXPORT_CSV", "./aspx_index.csv")),
        "SEARCH_FALLBACK": bool_from_env(env("SEARCH_FALLBACK"), default=True),
    }

    if site_url_value and isinstance(site_url_value, str):
        parts = urlsplit(site_url_value)
        if parts.netloc and not config["SPO_HOSTNAME"]:
            config["SPO_HOSTNAME"] = parts.netloc
        if parts.path:
            normalized_path = parts.path or "/"
            if not normalized_path.startswith("/"):
                normalized_path = f"/{normalized_path}"
            if config["SPO_SITE_PATH"] == "/":
                config["SPO_SITE_PATH"] = normalized_path

    missing = [
        key
        for key in ("AZURE_TENANT_ID", "AZURE_CLIENT_ID", "SPO_HOSTNAME")
        if not config[key]
    ]
    if missing:
        raise ValueError(
            "Missing required configuration values: " + ", ".join(missing)
        )

    if config["AUTH_MODE"] == "application" and not config["AZURE_CLIENT_SECRET"]:
        raise ValueError(
            "AZURE_CLIENT_SECRET must be provided for application auth mode."
        )

    return config


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------


def fail(message: str, *, code: int = 1) -> None:
    """Print an error message and exit."""

    print(f"[ERROR] {message}", file=sys.stderr)
    sys.exit(code)


def http_get_json(session: requests.Session, url: str, params=None, retries: int = 6):
    for attempt in range(retries):
        try:
            response = session.get(url, params=params)
        except requests.exceptions.SSLError as exc:  # pragma: no cover - network
            fail(f"SSL Error while GET {url}: {exc}")

        if response.ok:
            try:
                return response.json()
            except Exception:
                fail(f"Invalid JSON from {url}")

        if response.status_code in RETRY_STATUS:
            time.sleep(min(2**attempt, 30))
            continue

        fail(
            f"GET {url} failed: {response.status_code} {response.text[:300]}"
        )

    fail(f"GET {url} failed after retries.")


def http_post_json(
    session: requests.Session, url: str, body: dict, retries: int = 6
):
    for attempt in range(retries):
        try:
            response = session.post(url, json=body)
        except requests.exceptions.SSLError as exc:  # pragma: no cover - network
            fail(f"SSL Error while POST {url}: {exc}")

        if response.ok:
            try:
                return response.json()
            except Exception:
                fail(f"Invalid JSON from {url}")

        if response.status_code in RETRY_STATUS:
            time.sleep(min(2**attempt, 30))
            continue

        fail(
            f"POST {url} failed: {response.status_code} {response.text[:300]}"
        )

    fail(f"POST {url} failed after retries.")


def http_get_stream(
    session: requests.Session, url: str, retries: int = 3
) -> bytes:
    for attempt in range(retries):
        try:
            response = session.get(url, stream=True)
        except requests.exceptions.SSLError as exc:  # pragma: no cover - network
            fail(f"SSL Error while DOWNLOAD {url}: {exc}")

        if response.ok:
            return response.content

        if response.status_code in RETRY_STATUS:
            time.sleep(min(2**attempt, 30))
            continue

        fail(
            f"DOWNLOAD {url} failed: {response.status_code} {response.text[:300]}"
        )

    fail(f"DOWNLOAD {url} failed after retries.")


def is_folder(item: dict) -> bool:
    return "folder" in item or item.get("fields", {}).get("FSObjType") == 1


def is_aspx(item: dict) -> bool:
    name = (
        item.get("name")
        or item.get("fields", {}).get("FileLeafRef")
        or ""
    ).lower()
    return name.endswith(".aspx")


# ---------------------------------------------------------------------------
# Authentication helpers
# ---------------------------------------------------------------------------


def get_session(config: Dict[str, object]) -> requests.Session:
    tenant = str(config["AZURE_TENANT_ID"]).strip()
    client = str(config["AZURE_CLIENT_ID"]).strip()
    secret = str(config["AZURE_CLIENT_SECRET"] or "").strip()
    mode = str(config["AUTH_MODE"]).lower()

    print("🚀 Starting Microsoft Graph authentication...")

    if mode == "application":
        app = ConfidentialClientApplication(
            client,
            authority=f"https://login.microsoftonline.com/{tenant}",
            client_credential=secret,
        )
        token = app.acquire_token_for_client(GRAPH_SCOPE_APP)
    else:
        pca = PublicClientApplication(
            client, authority=f"https://login.microsoftonline.com/{tenant}"
        )
        flow = pca.initiate_device_flow(scopes=GRAPH_SCOPE_USER)
        if "user_code" not in flow:
            fail(str(flow))

        print(
            "\n=== DEVICE LOGIN ===\n"
            f"Go to: {flow['verification_uri']}\n"
            f"Code: {flow['user_code']}\n"
        )
        token = pca.acquire_token_by_device_flow(flow)

    if "access_token" not in token:
        fail(f"Auth failed: {token}")

    print("✅ Authentication successful!")

    session = requests.Session()
    session.verify = True
    session.headers.update({"Authorization": f"Bearer {token['access_token']}"})
    return session


# ---------------------------------------------------------------------------
# Microsoft Graph helper functions
# ---------------------------------------------------------------------------


def resolve_site_id(
    session: requests.Session, hostname: str, site_path: str
) -> str:
    print("🔍 Resolving site ID...")
    data = http_get_json(session, f"{GRAPH_BASE}/sites/{hostname}:{site_path}")
    site_id = data.get("id")
    if not site_id:
        fail(f"Could not resolve site: {hostname}:{site_path}")

    print(f"✅ Site resolved: {site_id}")
    return site_id


def list_drives(session: requests.Session, site_id: str) -> List[dict]:
    print("📦 Listing drives...")
    data = http_get_json(session, f"{GRAPH_BASE}/sites/{site_id}/drives")
    return data.get("value", [])


def list_lists(session: requests.Session, site_id: str) -> List[dict]:
    print("📋 Listing lists...")
    url = f"{GRAPH_BASE}/sites/{site_id}/lists?$select=id,displayName,webUrl"
    data = http_get_json(session, url)
    return data.get("value", [])


def get_all_sources(session: requests.Session, site_id: str) -> List[dict]:
    drives = list_drives(session, site_id)
    lists = list_lists(session, site_id)

    for lst in lists:
        name = (lst.get("displayName") or "").lower()
        if "page" in name or name in {"site pages", "pages"}:
            drives.append(
                {
                    "id": lst["id"],
                    "name": lst["displayName"],
                    "webUrl": lst.get("webUrl"),
                    "isList": True,
                }
            )

    if not any(
        "site" in drive.get("name", "").lower()
        and "page" in drive.get("name", "").lower()
        for drive in drives
    ):
        drives.append(
            {
                "id": "Site%20Pages",
                "name": "Site Pages",
                "webUrl": None,
                "isList": True,
            }
        )

    print(f"🔍 Content sources found: {[d['name'] for d in drives]}")
    return drives


def list_children_paged(session: requests.Session, url: str) -> List[dict]:
    items: List[dict] = []
    next_url: Optional[str] = url + "?$top=200"
    while next_url:
        data = http_get_json(session, next_url)
        items.extend(data.get("value", []))
        next_url = data.get("@odata.nextLink")
    return items


# ---------------------------------------------------------------------------
# Search fallback helpers
# ---------------------------------------------------------------------------


def graph_search_aspx_in_site(
    session: requests.Session, site_id: str
) -> List[dict]:
    print("🔎 Running Graph Search for filetype:aspx within site...")
    url = f"{GRAPH_BASE}/search/query"
    body = {
        "requests": [
            {
                "entityTypes": ["driveItem", "listItem"],
                "query": {"queryString": "filetype:aspx"},
                "fields": [
                    "name",
                    "webUrl",
                    "size",
                    "createdDateTime",
                    "lastModifiedDateTime",
                ],
                "from": 0,
                "size": 200,
                "sharePointOneDriveOptions": {
                    "includeContainers": [{"type": "site", "id": site_id}]
                },
            }
        ]
    }

    response = http_post_json(session, url, body)
    containers = (response.get("value") or [{}])[0].get("hitsContainers", [])

    results: List[dict] = []
    for container in containers:
        for hit in container.get("hits", []):
            resource = hit.get("resource", {}) or {}
            name = (resource.get("name") or "").lower()
            if not name.endswith(".aspx"):
                continue
            results.append(
                {
                    "drive": "Search",
                    "relPath": resource.get("name") or "",
                    "name": resource.get("name") or "",
                    "webUrl": resource.get("webUrl"),
                    "size": resource.get("size"),
                    "lastModifiedDateTime": resource.get("lastModifiedDateTime"),
                    "createdDateTime": resource.get("createdDateTime"),
                }
            )

    return results


def try_download_from_weburl(
    session: requests.Session, web_url: str, dest_path: pathlib.Path
) -> bool:
    try:
        response = session.get(web_url, stream=True)
        if response.ok and int(response.headers.get("Content-Length", "0") or 0) > 0:
            dest_path.parent.mkdir(parents=True, exist_ok=True)
            dest_path.write_bytes(response.content)
            return True
    except Exception:  # pragma: no cover - best effort helper
        return False
    return False


# ---------------------------------------------------------------------------
# Crawler implementation
# ---------------------------------------------------------------------------


def crawl_recursive(
    session: requests.Session,
    parent_url: str,
    rel_path: str,
    source_name: str,
    *,
    config: Dict[str, object],
    is_list: bool = False,
) -> List[dict]:
    collected: List[dict] = []
    children = list_children_paged(session, parent_url)
    for item in children:
        name = item.get("name") or item.get("fields", {}).get("FileLeafRef", "")
        if not name:
            continue

        relative_out = (
            (pathlib.PurePosixPath(rel_path) / name).as_posix()
            if rel_path
            else name
        )

        if is_folder(item):
            if "id" in item:
                base = parent_url.split("/items")[0]
                next_url = f"{base}/items/{item['id']}/children"
                collected.extend(
                    crawl_recursive(
                        session,
                        next_url,
                        relative_out,
                        source_name,
                        config=config,
                        is_list=is_list,
                    )
                )
            continue

        if not is_aspx(item):
            continue

        metadata = {
            "drive": source_name,
            "relPath": relative_out,
            "name": name,
            "webUrl": item.get("webUrl") or item.get("fields", {}).get("FileRef"),
            "size": item.get("size", ""),
            "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
            "createdDateTime": item.get("createdDateTime", ""),
        }
        collected.append(metadata)
        print(f"📄 Found: {metadata['relPath']}")

        download_enabled = bool(config["DOWNLOAD_ASPX"])
        if not download_enabled:
            continue

        destination = pathlib.Path(config["DOWNLOAD_DIR"]) / source_name / metadata["relPath"]
        destination.parent.mkdir(parents=True, exist_ok=True)

        try:
            if is_list:
                if metadata["webUrl"] and try_download_from_weburl(
                    session, metadata["webUrl"], destination
                ):
                    print(f"   ✅ Downloaded (webUrl) → {destination}")
                else:
                    print(
                        f"   ℹ️ Metadata only (list item): {metadata.get('webUrl', 'n/a')}"
                    )
            else:
                drive_id = parent_url.split("/drives/")[1].split("/")[0]
                content_url = (
                    f"{GRAPH_BASE}/drives/{drive_id}/items/{item['id']}/content"
                )
                bytes_data = http_get_stream(session, content_url)
                destination.write_bytes(bytes_data)
                print(f"   ✅ Downloaded → {destination}")
        except Exception as exc:  # pragma: no cover - network
            print(f"   ⚠️ Download failed: {exc}")

    return collected


def crawl_all_content(
    session: requests.Session,
    site_id: str,
    sources: Iterable[dict],
    *,
    config: Dict[str, object],
) -> List[dict]:
    results: List[dict] = []
    for source in sources:
        name = source.get("name", "Unnamed")
        print(f"\n📂 Scanning: {name}")
        if source.get("isList"):
            base_url = f"{GRAPH_BASE}/sites/{site_id}/lists/{source['id']}/items"
            results.extend(
                crawl_recursive(
                    session,
                    base_url,
                    "",
                    name,
                    config=config,
                    is_list=True,
                )
            )
        else:
            base_url = f"{GRAPH_BASE}/drives/{source['id']}/root/children"
            results.extend(
                crawl_recursive(
                    session,
                    base_url,
                    "",
                    name,
                    config=config,
                    is_list=False,
                )
            )

    return results


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def main(config: Dict[str, object]) -> None:
    print("🚀 Starting ASPX crawler...")
    session = get_session(config)
    site_id = resolve_site_id(
        session, str(config["SPO_HOSTNAME"]), str(config["SPO_SITE_PATH"])
    )
    sources = get_all_sources(session, site_id)
    results = crawl_all_content(session, site_id, sources, config=config)

    print(f"\n✅ Total .aspx files found by traversal: {len(results)}")

    if config["SEARCH_FALLBACK"] and len(results) == 0:
        print("🔎 No .aspx found by traversal — trying Graph Search fallback...")
        hits = graph_search_aspx_in_site(session, site_id)
        print(f"🔎 Graph Search found {len(hits)} .aspx candidates.")
        if config["DOWNLOAD_ASPX"]:
            for hit in hits:
                web_url = hit.get("webUrl")
                if not web_url:
                    continue
                destination = (
                    pathlib.Path(config["DOWNLOAD_DIR"]) / "Search" / hit["name"]
                )
                if try_download_from_weburl(session, web_url, destination):
                    print(f"   ✅ Downloaded (search) → {destination}")
                else:
                    print(f"   ⚠️ Could not download (search): {web_url}")
        results.extend(hits)

    print(f"\n✅ Total .aspx files reported: {len(results)}")

    export_csv: pathlib.Path = config["EXPORT_CSV"]
    export_csv.parent.mkdir(parents=True, exist_ok=True)
    with export_csv.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "drive",
                "name",
                "relPath",
                "webUrl",
                "size",
                "lastModifiedDateTime",
                "createdDateTime",
            ],
        )
        writer.writeheader()
        writer.writerows(results)

    print(f"📄 CSV saved to {export_csv}")
    print("✅ Crawl completed successfully.")


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="SharePoint ASPX crawler")
    parser.add_argument(
        "--config",
        help=(
            "Path to a JSON configuration file."
            " Values inside are merged with environment variables."
        ),
    )
    return parser.parse_args(argv)


if __name__ == "__main__":
    args = parse_args()
    try:
        config = load_config(config_path=args.config)
    except ValueError as exc:
        fail(str(exc))

    try:
        main(config)
    except KeyboardInterrupt:
        print("\n❌ Interrupted by user.")
        sys.exit(130)
