import base64
import io
import os
import posixpath
import time
import traceback
import heapq
from urllib.parse import unquote, urlparse, parse_qs
from datetime import datetime, timezone
from typing import Any

import requests
from openpyxl import load_workbook
from mcp.server.fastmcp import FastMCP

YANDEX_TOKEN = os.environ["YANDEX_DISK_TOKEN"].strip()
PORT = int(os.environ.get("PORT", "10000"))
MCP_SECRET = os.environ["MCP_SECRET"].strip()
if not YANDEX_TOKEN:
    raise RuntimeError("YANDEX_DISK_TOKEN is empty")
if not MCP_SECRET:
    raise RuntimeError("MCP_SECRET is empty")
YANDEX_API = "https://cloud-api.yandex.net/v1/disk"
APP_VERSION = "v7.3-fast-folder-20260830"

mcp = FastMCP(
    "Yandex Disk — Я Мебель v7.3",
    instructions=(
        "Read/write access to the user's Yandex Disk. "
        "Use list_folder/search_files to browse; read_excel/read_text_file to inspect; "
        "use write tools only when the user asks to modify files. "
        "Destructive operations should keep backups when available; delete_file defaults to Trash."
    ),
    host="0.0.0.0",
    port=PORT,
    stateless_http=True,
    json_response=True,
    streamable_http_path=f"/mcp/{MCP_SECRET}",
)

def yandex_headers() -> dict[str, str]:
    return {"Authorization": f"OAuth {YANDEX_TOKEN}"}

def _raise(r: requests.Response) -> None:
    if not r.ok:
        raise RuntimeError(f"Yandex Disk API error {r.status_code}: {r.text[:1000]}")

def _normalize_path(path: str) -> str:
    if not path:
        return "/"
    if path.startswith("disk:"):
        path = path[5:]
    path = "/" + path.lstrip("/")
    return posixpath.normpath(path)

def _parent(path: str) -> str:
    p = _normalize_path(path)
    parent = posixpath.dirname(p)
    return parent if parent else "/"

def _basename(path: str) -> str:
    return posixpath.basename(_normalize_path(path))

def _join(folder: str, name: str) -> str:
    return _normalize_path(posixpath.join(_normalize_path(folder), name))

def _resource(path: str, fields: str | None = None) -> dict[str, Any]:
    params: dict[str, Any] = {"path": _normalize_path(path)}
    if fields:
        params["fields"] = fields
    r = requests.get(f"{YANDEX_API}/resources", headers=yandex_headers(), params=params, timeout=30)
    _raise(r)
    return r.json()

def _resource_type(path: str) -> str | None:
    r = requests.get(
        f"{YANDEX_API}/resources",
        headers=yandex_headers(),
        params={"path": _normalize_path(path), "fields": "type"},
        timeout=30,
    )
    if r.status_code == 404:
        return None
    _raise(r)
    return r.json().get("type")

def _resource_exists(path: str) -> bool:
    return _resource_type(path) is not None

def _download_bytes(path: str) -> bytes:
    r = requests.get(
        f"{YANDEX_API}/resources/download",
        headers=yandex_headers(),
        params={"path": _normalize_path(path)},
        timeout=30,
    )
    _raise(r)
    href = r.json()["href"]
    d = requests.get(href, timeout=120)
    _raise(d)
    return d.content

def _upload_bytes(path: str, content: bytes, overwrite: bool = False) -> dict[str, Any]:
    path = _normalize_path(path)

    # 1) Ask Yandex Disk for a temporary upload URL.
    r = requests.get(
        f"{YANDEX_API}/resources/upload",
        headers=yandex_headers(),
        params={"path": path, "overwrite": str(bool(overwrite)).lower()},
        timeout=30,
    )
    _raise(r)

    upload_info = r.json()
    href = upload_info.get("href")
    method = str(upload_info.get("method") or "PUT").upper()

    if not href:
        raise RuntimeError(f"Yandex Disk did not return an upload URL for {path}")

    # 2) Upload the raw bytes to the temporary URL.
    upload_headers = {
        "Content-Type": "application/octet-stream",
        "Content-Length": str(len(content)),
    }

    if method == "PUT":
        u = requests.put(
            href,
            data=io.BytesIO(content),
            headers=upload_headers,
            timeout=180,
            allow_redirects=True,
        )
    elif method == "POST":
        u = requests.post(
            href,
            data=io.BytesIO(content),
            headers=upload_headers,
            timeout=180,
            allow_redirects=True,
        )
    else:
        raise RuntimeError(f"Unsupported upload method from Yandex Disk: {method}")

    _raise(u)

    # 3) Verify that the file actually appeared on Disk.
    # Small files normally appear immediately, but we retry briefly.
    verified = False
    last_error = None
    for _ in range(10):
        try:
            if _resource_exists(path):
                verified = True
                break
        except Exception as exc:
            last_error = str(exc)
        time.sleep(0.5)

    if not verified:
        raise RuntimeError(
            f"Upload request succeeded, but file was not found on Yandex Disk: {path}. "
            f"Upload status={u.status_code}, response={u.text[:500]!r}, "
            f"last_check_error={last_error!r}"
        )

    return {
        "ok": True,
        "path": path,
        "bytes_uploaded": len(content),
        "overwrite": overwrite,
        "upload_status": u.status_code,
        "verified": True,
    }

def _wait_operation(href: str | None, timeout_seconds: int = 60) -> dict[str, Any]:
    if not href:
        return {"status": "success"}
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        r = requests.get(href, headers=yandex_headers(), timeout=30)
        _raise(r)
        data = r.json()
        if data.get("status") in ("success", "failed"):
            return data
        time.sleep(1)
    return {"status": "in-progress", "href": href}

def _operation_result(r: requests.Response) -> dict[str, Any]:
    _raise(r)
    if r.status_code == 202:
        return _wait_operation(r.json().get("href"))
    if not r.content:
        return {"status": "success", "http_status": r.status_code}
    try:
        payload = r.json()
    except Exception:
        return {"status": "success", "http_status": r.status_code}
    if isinstance(payload, dict) and payload.get("status") in ("success", "failed", "in-progress"):
        return payload
    return {"status": "success", "http_status": r.status_code, "response": payload}

def _timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_%f")

def _backup_path(path: str) -> str:
    path = _normalize_path(path)
    folder = _parent(path)
    name = _basename(path)
    root, ext = posixpath.splitext(name)
    return _join(folder, f"{root}_backup_{_timestamp()}{ext}")

def _copy_internal(source_path: str, destination_path: str, overwrite: bool = False) -> dict[str, Any]:
    source_path = _normalize_path(source_path)
    destination_path = _normalize_path(destination_path)
    r = requests.post(
        f"{YANDEX_API}/resources/copy",
        headers=yandex_headers(),
        params={"from": source_path, "path": destination_path, "overwrite": str(bool(overwrite)).lower()},
        timeout=30,
    )
    operation = _operation_result(r)
    destination_exists = False
    if operation.get("status") == "success":
        for _ in range(10):
            destination_exists = _resource_exists(destination_path)
            if destination_exists:
                break
            time.sleep(0.4)
    return {
        "ok": operation.get("status") == "success" and destination_exists,
        "source": source_path,
        "destination": destination_path,
        "verified": destination_exists,
        "operation": operation,
    }

def _backup_if_requested(path: str, make_backup: bool) -> str | None:
    path = _normalize_path(path)
    if not make_backup or not _resource_exists(path):
        return None
    backup = _backup_path(path)
    result = _copy_internal(path, backup, overwrite=False)
    if not result.get("ok"):
        raise RuntimeError(f"Backup failed; original file was not modified: {result}")
    return backup

def _save_workbook(path: str, workbook, make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    backup = _backup_if_requested(path, make_backup)
    out = io.BytesIO()
    workbook.save(out)
    result = _upload_bytes(path, out.getvalue(), overwrite=True)
    result["backup_path"] = backup
    return result

def _load_workbook_from_disk(path: str):
    normalized = _normalize_path(path)
    keep_vba = normalized.lower().endswith((".xlsm", ".xltm"))
    return load_workbook(io.BytesIO(_download_bytes(normalized)), keep_vba=keep_vba)

def _private_path_from_yandex_url(value: str) -> str | None:
    """Extract a private Disk path from common Yandex Disk web-client URLs.

    Supported examples include /client/disk/<path> and URLs with a path query
    parameter. Public share links (/d/... or yadi.sk) intentionally return None
    and are handled through the public-resources endpoint instead.
    """
    try:
        parsed = urlparse(value)
    except Exception:
        return None
    host = (parsed.netloc or "").casefold()
    if not host or ("yandex." not in host and "yadi.sk" not in host):
        return None
    qs = parse_qs(parsed.query)
    for key in ("path", "dir"):
        vals = qs.get(key)
        if vals and vals[0]:
            return _normalize_path(unquote(vals[0]))
    marker = "/client/disk/"
    if marker in parsed.path:
        tail = parsed.path.split(marker, 1)[1]
        return _normalize_path(unquote(tail))
    if parsed.path.rstrip("/") == "/client/disk":
        return "/"
    return None


def _looks_like_public_yandex_url(value: str) -> bool:
    try:
        parsed = urlparse(value)
    except Exception:
        return False
    host = (parsed.netloc or "").casefold()
    path = parsed.path or ""
    return (
        host.endswith("yadi.sk")
        or ("disk.yandex." in host and (path.startswith("/d/") or path.startswith("/i/")))
    )


def _normalize_folder_item(i: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": i.get("name"),
        "type": i.get("type"),
        "path": i.get("path"),
        "size": i.get("size"),
        "mime_type": i.get("mime_type"),
        "modified": i.get("modified"),
        "created": i.get("created"),
    }


def _list_private_folder_page(path: str, offset: int = 0, limit: int = 200) -> dict[str, Any]:
    path = _normalize_path(path)
    limit = min(max(int(limit), 1), 500)
    offset = max(int(offset), 0)
    fields = (
        "name,type,path,_embedded.total,_embedded.limit,_embedded.offset,"
        "_embedded.items.name,_embedded.items.type,_embedded.items.path,"
        "_embedded.items.size,_embedded.items.mime_type,_embedded.items.modified,"
        "_embedded.items.created"
    )
    r = requests.get(
        f"{YANDEX_API}/resources",
        headers=yandex_headers(),
        params={"path": path, "limit": limit, "offset": offset, "fields": fields},
        timeout=10,
    )
    _raise(r)
    data = r.json()
    embedded = data.get("_embedded", {}) or {}
    batch = embedded.get("items", []) or []
    total = int(embedded.get("total", len(batch)) or 0)
    next_offset = offset + len(batch)
    return {
        "path": path,
        "items": [_normalize_folder_item(i) for i in batch],
        "offset": offset,
        "limit": limit,
        "total": total,
        "has_more": next_offset < total,
        "next_offset": next_offset if next_offset < total else None,
    }


def _list_public_folder_page(public_url: str, offset: int = 0, limit: int = 200) -> dict[str, Any]:
    limit = min(max(int(limit), 1), 500)
    offset = max(int(offset), 0)
    fields = (
        "name,type,path,public_url,_embedded.total,_embedded.limit,_embedded.offset,"
        "_embedded.items.name,_embedded.items.type,_embedded.items.path,"
        "_embedded.items.size,_embedded.items.mime_type,_embedded.items.modified,"
        "_embedded.items.created"
    )
    r = requests.get(
        f"{YANDEX_API}/public/resources",
        params={"public_key": public_url, "limit": limit, "offset": offset, "fields": fields},
        timeout=10,
    )
    _raise(r)
    data = r.json()
    embedded = data.get("_embedded", {}) or {}
    batch = embedded.get("items", []) or []
    total = int(embedded.get("total", len(batch)) or 0)
    next_offset = offset + len(batch)
    return {
        "public_url": public_url,
        "name": data.get("name"),
        "items": [_normalize_folder_item(i) for i in batch],
        "offset": offset,
        "limit": limit,
        "total": total,
        "has_more": next_offset < total,
        "next_offset": next_offset if next_offset < total else None,
    }


def _list_folder_internal(path: str = "/", offset: int = 0, limit: int = 200) -> dict[str, Any]:
    """Return one lightweight page instead of exhaustively reading the folder.

    This is intentionally bounded so large directories return quickly. It also
    accepts common Yandex Disk web links: private web-client URLs are converted
    to a Disk path, while public share URLs use /public/resources.
    """
    raw = str(path or "/").strip()
    if raw.startswith("http://") or raw.startswith("https://"):
        private_path = _private_path_from_yandex_url(raw)
        if private_path is not None:
            return _list_private_folder_page(private_path, offset=offset, limit=limit)
        if _looks_like_public_yandex_url(raw):
            return _list_public_folder_page(raw, offset=offset, limit=limit)
        raise ValueError("Unsupported Yandex Disk URL format")
    return _list_private_folder_page(raw, offset=offset, limit=limit)


@mcp.tool()
def list_folder(path: str = "/", offset: int = 0, limit: int = 200) -> dict[str, Any]:
    """List one fast page of a Yandex Disk folder.

    Accepts a Disk path or common Yandex Disk folder URL. Returns at most `limit`
    items plus total/has_more/next_offset, so large folders do not block the MCP
    request while every page is fetched.
    """
    return _list_folder_internal(path, offset=offset, limit=limit)

@mcp.tool()
def get_file_info(path: str) -> dict[str, Any]:
    return _resource(path, fields="name,type,path,size,mime_type,created,modified,md5,sha256,revision")

@mcp.tool()
def file_exists(path: str) -> bool:
    return _resource_type(path) == "file"

@mcp.tool()
def folder_exists(path: str) -> bool:
    return _resource_type(path) == "dir"

def _search_folder_page(path: str, offset: int = 0, limit: int = 250) -> tuple[list[dict[str, Any]], int]:
    """Fetch one lightweight folder page for search.

    Unlike _list_folder_internal(), this intentionally does not read every page of a
    large directory before search can continue. That keeps MCP calls responsive.
    """
    path = _normalize_path(path)
    fields = (
        "_embedded.items.name,_embedded.items.type,_embedded.items.path,"
        "_embedded.items.size,_embedded.items.mime_type,_embedded.items.modified,"
        "_embedded.items.created,_embedded.total"
    )
    r = requests.get(
        f"{YANDEX_API}/resources",
        headers=yandex_headers(),
        params={"path": path, "limit": limit, "offset": offset, "fields": fields},
        timeout=10,
    )
    _raise(r)
    embedded = r.json().get("_embedded", {})
    items = embedded.get("items", []) or []
    total = int(embedded.get("total", len(items)) or 0)
    normalized_items: list[dict[str, Any]] = []
    for i in items:
        normalized_items.append({
            "name": i.get("name"),
            "type": i.get("type"),
            "path": i.get("path"),
            "size": i.get("size"),
            "mime_type": i.get("mime_type"),
            "modified": i.get("modified"),
            "created": i.get("created"),
        })
    return normalized_items, total


def _search_priority(name: str, query_norm: str, tokens: list[str], depth: int) -> int:
    """Lower number means search this directory sooner."""
    name_norm = name.casefold()
    if name_norm == query_norm:
        return -10000
    if query_norm in name_norm:
        return -5000 + depth
    token_hits = sum(1 for token in tokens if token in name_norm)
    if token_hits:
        return -1000 * token_hits + depth
    return 100 + depth


@mcp.tool()
def search_files(query: str, start_path: str = "/", max_results: int = 100, max_depth: int = 8) -> list[dict[str, Any]]:
    """Fast bounded name search across Yandex Disk.

    The old implementation fully listed each visited directory before moving on,
    which could time out on a large Disk. This version:
    - returns immediately on an exact name match;
    - reads large directories page-by-page;
    - prioritizes folders whose names resemble the query;
    - uses a hard time budget so the MCP call returns instead of hanging.
    """
    query_norm = query.casefold().strip()
    if not query_norm:
        raise ValueError("query must not be empty")

    start_path = _normalize_path(start_path)
    max_results = min(max(int(max_results), 1), 1000)
    max_depth = min(max(int(max_depth), 0), 32)
    tokens = [part for part in query_norm.replace("_", " ").replace("-", " ").split() if len(part) >= 2]

    # Very cheap common case: the requested name is a direct child of start_path.
    direct_candidate = _join(start_path, query.strip())
    try:
        info = _resource(
            direct_candidate,
            fields="name,type,path,size,mime_type,created,modified",
        )
        if str(info.get("name") or "").casefold() == query_norm:
            return [{
                "name": info.get("name"),
                "type": info.get("type"),
                "path": info.get("path"),
                "size": info.get("size"),
                "mime_type": info.get("mime_type"),
                "modified": info.get("modified"),
                "created": info.get("created"),
            }]
    except Exception:
        pass

    # Stay well below the outer MCP/ChatGPT timeout. Returning a partial result is
    # preferable to making the whole tool call fail.
    deadline = time.monotonic() + 15.0
    page_limit = 250
    results: list[dict[str, Any]] = []
    seen_paths: set[str] = set()
    seen_result_paths: set[str] = set()
    sequence = 0

    # Heap item: (priority, depth, sequence, folder_path, offset)
    queue: list[tuple[int, int, int, str, int]] = [(0, 0, sequence, start_path, 0)]

    while queue and len(results) < max_results and time.monotonic() < deadline:
        _, depth, _, folder, offset = heapq.heappop(queue)
        page_key = f"{folder}#{offset}"
        if page_key in seen_paths:
            continue
        seen_paths.add(page_key)

        try:
            items, total = _search_folder_page(folder, offset=offset, limit=page_limit)
        except Exception:
            continue

        # Process matches first, so an exact match returns without walking the tree.
        for item in items:
            name = str(item.get("name") or "")
            name_norm = name.casefold()
            raw_path = item.get("path") or ""
            result_path = str(raw_path)

            if name_norm == query_norm:
                return [item]

            if query_norm in name_norm and result_path not in seen_result_paths:
                results.append(item)
                seen_result_paths.add(result_path)
                if len(results) >= max_results:
                    return results

        # Continue scanning a huge directory later rather than blocking on all pages.
        next_offset = offset + len(items)
        if items and next_offset < total:
            sequence += 1
            heapq.heappush(queue, (50 + depth, depth, sequence, folder, next_offset))

        if depth >= max_depth:
            continue

        # Prioritize promising directory names such as "Ростов" or "Филиал".
        for item in items:
            if item.get("type") != "dir":
                continue
            raw_path = item.get("path") or ""
            if isinstance(raw_path, str) and raw_path.startswith("disk:"):
                raw_path = raw_path[5:]
            child_path = _normalize_path(str(raw_path))
            name = str(item.get("name") or "")
            priority = _search_priority(name, query_norm, tokens, depth + 1)
            sequence += 1
            heapq.heappush(queue, (priority, depth + 1, sequence, child_path, 0))

    return results

@mcp.tool()
def get_file_download_url(path: str) -> dict[str, str]:
    path = _normalize_path(path)
    r = requests.get(f"{YANDEX_API}/resources/download", headers=yandex_headers(), params={"path": path}, timeout=30)
    _raise(r)
    return {"path": path, "download_url": r.json()["href"]}

@mcp.tool()
def read_text_file(path: str, max_chars: int = 50000, encoding: str = "utf-8") -> dict[str, Any]:
    text = _download_bytes(path).decode(encoding, errors="replace")
    return {"path": _normalize_path(path), "chars_total": len(text), "truncated": len(text) > max_chars, "text": text[:max_chars]}

@mcp.tool()
def read_excel(path: str, sheet: str = "", max_rows: int = 200) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = []
    for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx > max_rows:
            break
        rows.append(list(row))
    return {"path": _normalize_path(path), "sheets": wb.sheetnames, "selected_sheet": ws.title, "rows_returned": len(rows), "rows": rows}

@mcp.tool()
def read_excel_sheet_names(path: str) -> list[str]:
    return _load_workbook_from_disk(path).sheetnames

@mcp.tool()
def create_folder(path: str) -> dict[str, Any]:
    path = _normalize_path(path)
    r = requests.put(f"{YANDEX_API}/resources", headers=yandex_headers(), params={"path": path}, timeout=30)
    if r.status_code == 409:
        existing_type = _resource_type(path)
        if existing_type == "dir":
            return {"ok": True, "path": path, "already_exists": True}
        if existing_type == "file":
            raise RuntimeError(f"Cannot create folder because a file already exists at: {path}")
    _raise(r)
    return {"ok": True, "path": path, "already_exists": False}

@mcp.tool()
def copy_file(source_path: str, destination_path: str, overwrite: bool = False) -> dict[str, Any]:
    return _copy_internal(source_path, destination_path, overwrite)

def _move_diagnostic(source_path: str, destination_path: str, overwrite: bool = False, execute: bool = True) -> dict[str, Any]:
    """Run or inspect a Yandex Disk move without hiding any diagnostic detail.

    This helper deliberately avoids raising for Yandex API errors so the MCP client
    receives the HTTP status/body directly.
    """
    result: dict[str, Any] = {
        "ok": False,
        "bridge_version": APP_VERSION,
        "execute": execute,
        "source": None,
        "destination": None,
        "overwrite": bool(overwrite),
        "stages": [],
    }
    try:
        source_path = _normalize_path(source_path)
        destination_path = _normalize_path(destination_path)
        result["source"] = source_path
        result["destination"] = destination_path
        result["stages"].append({"stage": "normalized"})

        source_exists_before = _resource_exists(source_path)
        destination_exists_before = _resource_exists(destination_path)
        result["source_exists_before"] = source_exists_before
        result["destination_exists_before"] = destination_exists_before
        result["stages"].append({
            "stage": "precheck",
            "source_exists": source_exists_before,
            "destination_exists": destination_exists_before,
        })

        if not source_exists_before:
            result["error_type"] = "FileNotFoundError"
            result["error"] = f"Source does not exist: {source_path}"
            return result

        if source_path == destination_path:
            result.update({"ok": True, "already_at_destination": True, "verified": True})
            result["stages"].append({"stage": "no_op_same_path"})
            return result

        if destination_exists_before and not overwrite:
            result["error_type"] = "DestinationExistsError"
            result["error"] = f"Destination already exists: {destination_path}"
            return result

        params = {
            "from": source_path,
            "path": destination_path,
            "overwrite": str(bool(overwrite)).lower(),
        }
        result["request"] = {
            "method": "POST",
            "endpoint": "/resources/move",
            "params": params,
        }

        if not execute:
            result["ok"] = True
            result["stages"].append({"stage": "dry_run_ready"})
            return result

        r = requests.post(
            f"{YANDEX_API}/resources/move",
            headers=yandex_headers(),
            params=params,
            timeout=30,
        )
        body = (r.text or "")[:4000]
        result["yandex_http_status"] = r.status_code
        result["yandex_response_body"] = body
        result["stages"].append({"stage": "move_http_response", "status": r.status_code})

        parsed: Any = None
        if r.content:
            try:
                parsed = r.json()
            except Exception:
                parsed = None
        if parsed is not None:
            result["yandex_response_json"] = parsed

        if not r.ok:
            result["error_type"] = "YandexDiskHTTPError"
            result["error"] = f"Yandex Disk API returned HTTP {r.status_code}"
            return result

        operation: dict[str, Any] = {"status": "success"}
        if r.status_code == 202:
            href = parsed.get("href") if isinstance(parsed, dict) else None
            result["operation_href_present"] = bool(href)
            if href:
                deadline = time.time() + 30
                poll_count = 0
                while time.time() < deadline:
                    poll_count += 1
                    pr = requests.get(href, headers=yandex_headers(), timeout=30)
                    pbody = (pr.text or "")[:4000]
                    try:
                        pdata = pr.json() if pr.content else {}
                    except Exception:
                        pdata = {"raw_body": pbody}
                    operation = {
                        "http_status": pr.status_code,
                        "body": pdata,
                        "poll_count": poll_count,
                    }
                    status = pdata.get("status") if isinstance(pdata, dict) else None
                    if not pr.ok or status in ("success", "failed"):
                        break
                    time.sleep(0.5)
            else:
                operation = {"status": "unknown", "reason": "202 response without href"}
        elif isinstance(parsed, dict) and parsed:
            operation = parsed

        result["operation"] = operation
        result["stages"].append({"stage": "operation_checked"})

        destination_exists_after = False
        source_exists_after = True
        verify_attempts = 0
        for verify_attempts in range(1, 11):
            destination_exists_after = _resource_exists(destination_path)
            source_exists_after = _resource_exists(source_path)
            if destination_exists_after and not source_exists_after:
                break
            time.sleep(0.4)

        result["source_exists_after"] = source_exists_after
        result["destination_exists_after"] = destination_exists_after
        result["verification_attempts"] = verify_attempts
        result["verified"] = destination_exists_after and not source_exists_after
        result["stages"].append({
            "stage": "postcheck",
            "source_exists": source_exists_after,
            "destination_exists": destination_exists_after,
            "attempts": verify_attempts,
        })

        if result["verified"]:
            result["ok"] = True
        else:
            result["error_type"] = "MoveVerificationError"
            result["error"] = "Yandex request completed but the expected final file state was not observed after verification retries."
        return result
    except Exception as exc:
        result["error_type"] = type(exc).__name__
        result["error"] = str(exc)
        result["traceback_tail"] = traceback.format_exc().splitlines()[-12:]
        return result


def _move_internal(source_path: str, destination_path: str, overwrite: bool = False) -> dict[str, Any]:
    result = _move_diagnostic(source_path, destination_path, overwrite, execute=True)
    if not result.get("ok"):
        raise RuntimeError(f"Move failed: {result}")
    return result


@mcp.tool()
def self_check() -> dict[str, Any]:
    """Non-destructive bridge check: version, configuration and Yandex Disk API access."""
    result: dict[str, Any] = {
        "ok": False,
        "bridge_version": APP_VERSION,
        "server": "Yandex Disk — Я Мебель v7.3",
        "mcp_secret_configured": bool(MCP_SECRET),
        "yandex_token_configured": bool(YANDEX_TOKEN),
    }
    try:
        r = requests.get(
            f"{YANDEX_API}/resources",
            headers=yandex_headers(),
            params={"path": "/", "limit": 1, "fields": "name,type,path"},
            timeout=30,
        )
        result["yandex_http_status"] = r.status_code
        result["ok"] = bool(r.ok)
        if not r.ok:
            result["error"] = (r.text or "")[:1000]
        return result
    except Exception as exc:
        result["error_type"] = type(exc).__name__
        result["error"] = str(exc)
        return result


@mcp.tool()
def bridge_version() -> dict[str, Any]:
    return {"ok": True, "bridge_version": APP_VERSION}


@mcp.tool()
def diagnose_move(source_path: str, destination_path: str, overwrite: bool = False, execute: bool = False) -> dict[str, Any]:
    """Return move diagnostics directly to ChatGPT. Set execute=true to perform the move."""
    return _move_diagnostic(source_path, destination_path, overwrite, execute=execute)


@mcp.tool()
def move_file(source_path: str, destination_path: str, overwrite: bool = False) -> dict[str, Any]:
    detail = _move_diagnostic(source_path, destination_path, overwrite, execute=True)
    result = {
        "ok": bool(detail.get("ok")),
        "bridge_version": APP_VERSION,
        "source": detail.get("source"),
        "destination": detail.get("destination"),
        "overwrite": bool(overwrite),
        "verified": bool(detail.get("verified")),
        "yandex_http_status": detail.get("yandex_http_status"),
        "operation": detail.get("operation"),
    }
    if not result["ok"]:
        result["error_type"] = detail.get("error_type", "MoveFailed")
        result["error"] = detail.get("error", "Move failed")
    return result


@mcp.tool()
def rename_file(path: str, new_name: str, overwrite: bool = False) -> dict[str, Any]:
    result: dict[str, Any] = {
        "ok": False,
        "bridge_version": APP_VERSION,
        "path": path,
        "new_name": new_name,
        "overwrite": bool(overwrite),
    }
    try:
        normalized_path = _normalize_path(path)
        clean_name = new_name.strip()
        result["path"] = normalized_path
        result["new_name"] = clean_name
        if not clean_name or clean_name in (".", "..") or "/" in clean_name or "\\" in clean_name:
            result["error_type"] = "ValueError"
            result["error"] = "new_name must be a safe basename only"
            return result

        destination_path = _join(_parent(normalized_path), clean_name)
        result["destination"] = destination_path
        move_result = _move_diagnostic(normalized_path, destination_path, overwrite, execute=True)
        result["ok"] = bool(move_result.get("ok"))
        result["verified"] = bool(move_result.get("verified"))
        result["yandex_http_status"] = move_result.get("yandex_http_status")
        result["operation"] = move_result.get("operation")
        if not result["ok"]:
            result["error_type"] = move_result.get("error_type", "RenameFailed")
            result["error"] = move_result.get("error", "Rename failed")
        return result
    except Exception as exc:
        result["error_type"] = type(exc).__name__
        result["error"] = str(exc)
        result["traceback_tail"] = traceback.format_exc().splitlines()[-12:]
        return result


@mcp.tool()
def backup_file(path: str) -> dict[str, Any]:
    destination = _backup_path(path)
    result = _copy_internal(path, destination, False)
    result["backup_path"] = destination
    return result

@mcp.tool()
def delete_file(path: str, permanently: bool = False) -> dict[str, Any]:
    path = _normalize_path(path)
    r = requests.delete(
        f"{YANDEX_API}/resources",
        headers=yandex_headers(),
        params={"path": path, "permanently": str(permanently).lower()},
        timeout=30,
    )
    op = _operation_result(r)
    removed = False
    if op.get("status") == "success":
        for _ in range(10):
            removed = not _resource_exists(path)
            if removed:
                break
            time.sleep(0.4)
    return {
        "ok": op.get("status") == "success" and removed,
        "path": path,
        "permanently": permanently,
        "verified": removed,
        "operation": op,
    }

@mcp.tool()
def list_trash(limit: int = 100) -> list[dict[str, Any]]:
    r = requests.get(f"{YANDEX_API}/trash/resources", headers=yandex_headers(), params={"limit": min(max(limit, 1), 1000)}, timeout=30)
    _raise(r)
    items = r.json().get("_embedded", {}).get("items", [])
    return [{"name": i.get("name"), "type": i.get("type"), "path": i.get("path"), "deleted": i.get("deleted"), "origin_path": i.get("origin_path")} for i in items]

@mcp.tool()
def restore_from_trash(trash_path: str, new_name: str = "", overwrite: bool = False) -> dict[str, Any]:
    meta_response = requests.get(
        f"{YANDEX_API}/trash/resources",
        headers=yandex_headers(),
        params={"path": trash_path, "fields": "name,path,origin_path"},
        timeout=30,
    )
    _raise(meta_response)
    meta = meta_response.json()
    origin_path = meta.get("origin_path")
    clean_name = new_name.strip()
    if clean_name and (clean_name in (".", "..") or "/" in clean_name or "\\" in clean_name):
        raise ValueError("new_name must be a safe basename only")

    params: dict[str, Any] = {"path": trash_path, "overwrite": str(bool(overwrite)).lower()}
    if clean_name:
        params["name"] = clean_name
    r = requests.put(f"{YANDEX_API}/trash/resources/restore", headers=yandex_headers(), params=params, timeout=30)
    op = _operation_result(r)

    expected_path = None
    verified = False
    if origin_path:
        expected_path = _normalize_path(origin_path)
        if clean_name:
            expected_path = _join(_parent(expected_path), clean_name)
        if op.get("status") == "success":
            for _ in range(10):
                verified = _resource_exists(expected_path)
                if verified:
                    break
                time.sleep(0.4)

    ok = op.get("status") == "success" and (verified if expected_path else True)
    return {
        "ok": ok,
        "trash_path": trash_path,
        "new_name": clean_name or None,
        "overwrite": bool(overwrite),
        "origin_path": origin_path,
        "restored_path": expected_path,
        "verified": verified if expected_path else None,
        "operation": op,
    }

@mcp.tool()
def upload_file(path: str, content_base64: str, overwrite: bool = False) -> dict[str, Any]:
    try:
        data = base64.b64decode(content_base64, validate=True)
    except Exception as exc:
        raise ValueError(f"Invalid base64 content: {exc}") from exc
    return _upload_bytes(path, data, overwrite)

@mcp.tool()
def overwrite_file(path: str, content_base64: str, make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    try:
        data = base64.b64decode(content_base64, validate=True)
    except Exception as exc:
        raise ValueError(f"Invalid base64 content: {exc}") from exc
    backup = _backup_if_requested(path, make_backup)
    result = _upload_bytes(path, data, True)
    result["backup_path"] = backup
    return result

@mcp.tool()
def write_text_file(path: str, text: str, overwrite: bool = False, encoding: str = "utf-8", make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    data = text.encode(encoding)
    backup = _backup_if_requested(path, make_backup) if overwrite else None
    result = _upload_bytes(path, data, overwrite)
    result["backup_path"] = backup
    return result

@mcp.tool()
def append_text_file(path: str, text: str, encoding: str = "utf-8", make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    existing = _download_bytes(path).decode(encoding, errors="replace")
    data = (existing + text).encode(encoding)
    backup = _backup_if_requested(path, make_backup)
    result = _upload_bytes(path, data, True)
    result["backup_path"] = backup
    return result

@mcp.tool()
def replace_text(path: str, old_text: str, new_text: str, replace_all: bool = True, encoding: str = "utf-8", make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    if old_text == "":
        raise ValueError("old_text must not be empty")
    existing = _download_bytes(path).decode(encoding, errors="replace")
    count = existing.count(old_text)
    if count == 0:
        return {"ok": False, "path": path, "replacements": 0, "message": "old_text not found"}
    updated = existing.replace(old_text, new_text) if replace_all else existing.replace(old_text, new_text, 1)
    data = updated.encode(encoding)
    backup = _backup_if_requested(path, make_backup)
    result = _upload_bytes(path, data, True)
    result.update({"backup_path": backup, "replacements": count if replace_all else 1})
    return result

@mcp.tool()
def update_excel_cell(path: str, sheet: str, cell: str, value: Any, make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet}")
    ws = wb[sheet]
    old_value = ws[cell].value
    ws[cell] = value
    result = _save_workbook(path, wb, make_backup)
    result.update({"sheet": sheet, "cell": cell, "old_value": old_value, "new_value": value})
    return result

def _update_excel_range_internal(path: str, sheet: str, start_cell: str, values: list[list[Any]], make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet}")
    ws = wb[sheet]
    anchor = ws[start_cell]
    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            ws.cell(row=anchor.row + r_idx, column=anchor.column + c_idx, value=value)
    result = _save_workbook(path, wb, make_backup)
    result.update({"sheet": sheet, "start_cell": start_cell, "rows_written": len(values), "max_columns_written": max((len(r) for r in values), default=0)})
    return result


@mcp.tool()
def update_excel_range(path: str, sheet: str, start_cell: str, values: list[list[Any]], make_backup: bool = True) -> dict[str, Any]:
    return _update_excel_range_internal(path, sheet, start_cell, values, make_backup)

@mcp.tool()
def append_excel_row(path: str, sheet: str, values: list[Any], make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if sheet not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet}")
    ws = wb[sheet]
    ws.append(values)
    row_number = ws.max_row
    result = _save_workbook(path, wb, make_backup)
    result.update({"sheet": sheet, "row_number": row_number, "values": values})
    return result

@mcp.tool()
def write_excel_table(path: str, sheet: str, start_cell: str, rows: list[list[Any]], make_backup: bool = True) -> dict[str, Any]:
    return _update_excel_range_internal(path, sheet, start_cell, rows, make_backup)

@mcp.tool()
def add_excel_sheet(path: str, sheet_name: str, index: int = -1, make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if sheet_name in wb.sheetnames:
        raise ValueError(f"Sheet already exists: {sheet_name}")
    if index < 0:
        wb.create_sheet(title=sheet_name)
    else:
        wb.create_sheet(title=sheet_name, index=index)
    result = _save_workbook(path, wb, make_backup)
    result.update({"sheet_added": sheet_name, "sheet_names": wb.sheetnames})
    return result

@mcp.tool()
def rename_excel_sheet(path: str, old_name: str, new_name: str, make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if old_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {old_name}")
    if new_name in wb.sheetnames:
        raise ValueError(f"Sheet already exists: {new_name}")
    wb[old_name].title = new_name
    result = _save_workbook(path, wb, make_backup)
    result.update({"old_name": old_name, "new_name": new_name, "sheet_names": wb.sheetnames})
    return result

@mcp.tool()
def delete_excel_sheet(path: str, sheet_name: str, make_backup: bool = True) -> dict[str, Any]:
    wb = _load_workbook_from_disk(path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet_name}")
    if len(wb.sheetnames) <= 1:
        raise ValueError("Cannot delete the only worksheet in the workbook")
    del wb[sheet_name]
    result = _save_workbook(path, wb, make_backup)
    result.update({"sheet_deleted": sheet_name, "sheet_names": wb.sheetnames})
    return result

if __name__ == "__main__":
    print(f"BOOTING {APP_VERSION} from {__file__}", flush=True)
    mcp.run(transport="streamable-http")
