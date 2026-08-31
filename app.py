import base64
import csv
import io
import json
import os
import posixpath
import time
import traceback
import heapq
import hashlib
import tempfile
import uuid
from pathlib import Path
from urllib.parse import unquote, urlparse, parse_qs
from datetime import datetime, timezone
from typing import Any

import requests
from openpyxl import load_workbook
from docx import Document
from pypdf import PdfReader
from pptx import Presentation
from PIL import Image as PILImage, ImageOps
try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    # HEIC/HEIF support is optional at runtime; common JPEG/PNG/etc. still work.
    pass
from mcp.server.fastmcp import FastMCP, Image as MCPImage

YANDEX_TOKEN = os.environ["YANDEX_DISK_TOKEN"].strip()
PORT = int(os.environ.get("PORT", "10000"))
MCP_SECRET = os.environ["MCP_SECRET"].strip()
if not YANDEX_TOKEN:
    raise RuntimeError("YANDEX_DISK_TOKEN is empty")
if not MCP_SECRET:
    raise RuntimeError("MCP_SECRET is empty")
YANDEX_API = "https://cloud-api.yandex.net/v1/disk"
APP_VERSION = "v7.6.3a-strict-base64-2k-fix-20260831"

mcp = FastMCP(
    "Yandex Disk — Я Мебель v7.6.3a strict-base64-2k fix",
    instructions=(
        "Read/write access to the user's Yandex Disk. "
        "Use list_folder/search_files to browse; read_file/read_docx/read_pdf/read_pptx/read_excel/read_text_file/read_image to inspect; "
        "use write tools only when the user asks to modify files. "
        "Destructive operations should keep backups when available; delete_file defaults to Trash. "
        "For generated or large binary files use begin_chunked_upload/upload_file_chunk_base64/finish_chunked_upload."
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

MAX_SERVER_DOWNLOAD_BYTES = int(os.environ.get("MAX_SERVER_DOWNLOAD_BYTES", str(100 * 1024 * 1024)))

def _download_bytes(path: str, max_bytes: int | None = None) -> bytes:
    """Download a Yandex Disk file on the Render server, not in ChatGPT."""
    normalized = _normalize_path(path)
    hard_limit = MAX_SERVER_DOWNLOAD_BYTES if max_bytes is None else min(max(int(max_bytes), 1), MAX_SERVER_DOWNLOAD_BYTES)
    r = requests.get(
        f"{YANDEX_API}/resources/download",
        headers=yandex_headers(),
        params={"path": normalized},
        timeout=30,
    )
    _raise(r)
    href = r.json().get("href")
    if not href:
        raise RuntimeError(f"Yandex Disk did not return a download URL for {normalized}")
    with requests.get(href, timeout=(15, 180), stream=True, allow_redirects=True) as d:
        _raise(d)
        length = d.headers.get("Content-Length")
        if length is not None:
            try:
                content_length = int(length)
            except (TypeError, ValueError):
                content_length = None
            if content_length is not None and content_length > hard_limit:
                raise ValueError(f"File is too large for server-side reading: {content_length} bytes; limit is {hard_limit} bytes")
        out = io.BytesIO()
        total = 0
        for chunk in d.iter_content(chunk_size=256 * 1024):
            if not chunk:
                continue
            total += len(chunk)
            if total > hard_limit:
                raise ValueError(f"File exceeded server-side reading limit while downloading: >{hard_limit} bytes")
            out.write(chunk)
        return out.getvalue()

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


# Chunked upload staging. Render's filesystem is ephemeral, so an in-progress
# session survives normal MCP calls on the same instance but not a service restart.
CHUNK_UPLOAD_DIR = Path(os.environ.get("CHUNK_UPLOAD_DIR", tempfile.gettempdir())) / "yandex_disk_bridge_uploads"
CHUNK_UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
MAX_CHUNK_UPLOAD_BYTES = int(os.environ.get("MAX_CHUNK_UPLOAD_BYTES", str(250 * 1024 * 1024)))
CHUNK_UPLOAD_TTL_SECONDS = int(os.environ.get("CHUNK_UPLOAD_TTL_SECONDS", str(6 * 60 * 60)))
# MCP request payloads can be truncated by clients/proxies before they reach the bridge.
# Keep upload chunks intentionally small. The environment can lower this value, but
# the bridge never accepts more than 16 KiB per chunk in v7.6.2.
MAX_CHUNK_BYTES = int(os.environ.get("MAX_CHUNK_BYTES", str(8_000)))
SAFE_MCP_CHUNK_CAP_BYTES = 16_000
RECOMMENDED_MCP_CHUNK_BYTES = 2048


def _upload_local_file(path: str, local_path: Path, overwrite: bool = False) -> dict[str, Any]:
    """Stream a local staged file to Yandex Disk without loading it all into RAM."""
    path = _normalize_path(path)
    local_path = Path(local_path)
    size = local_path.stat().st_size
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
    upload_headers = {"Content-Type": "application/octet-stream", "Content-Length": str(size)}
    with local_path.open("rb") as fh:
        if method == "PUT":
            u = requests.put(href, data=fh, headers=upload_headers, timeout=(15, 600), allow_redirects=True)
        elif method == "POST":
            u = requests.post(href, data=fh, headers=upload_headers, timeout=(15, 600), allow_redirects=True)
        else:
            raise RuntimeError(f"Unsupported upload method from Yandex Disk: {method}")
    _raise(u)
    verified = False
    remote_size = None
    last_error = None
    for _ in range(12):
        try:
            meta = _resource(path, fields="type,size")
            remote_size = meta.get("size")
            if meta.get("type") == "file" and (remote_size is None or int(remote_size) == size):
                verified = True
                break
        except Exception as exc:
            last_error = str(exc)
        time.sleep(0.5)
    if not verified:
        raise RuntimeError(
            f"Upload request completed but verification failed for {path}: local_size={size}, "
            f"remote_size={remote_size}, status={u.status_code}, last_error={last_error!r}"
        )
    return {
        "ok": True,
        "path": path,
        "bytes_uploaded": size,
        "overwrite": bool(overwrite),
        "upload_status": u.status_code,
        "verified": True,
        "remote_size": remote_size,
    }


def _upload_session_paths(upload_id: str) -> tuple[Path, Path]:
    try:
        normalized_id = str(uuid.UUID(str(upload_id)))
    except Exception as exc:
        raise ValueError("Invalid upload_id") from exc
    return CHUNK_UPLOAD_DIR / f"{normalized_id}.json", CHUNK_UPLOAD_DIR / f"{normalized_id}.part"


def _cleanup_stale_uploads() -> int:
    now = time.time()
    removed = 0
    for meta_path in CHUNK_UPLOAD_DIR.glob("*.json"):
        try:
            age = now - meta_path.stat().st_mtime
            if age <= CHUNK_UPLOAD_TTL_SECONDS:
                continue
            part_path = meta_path.with_suffix(".part")
            meta_path.unlink(missing_ok=True)
            part_path.unlink(missing_ok=True)
            removed += 1
        except Exception:
            continue
    return removed


def _load_upload_session(upload_id: str) -> tuple[dict[str, Any], Path, Path]:
    meta_path, part_path = _upload_session_paths(upload_id)
    if not meta_path.exists():
        raise ValueError("Upload session not found or expired")
    try:
        meta = json.loads(meta_path.read_text("utf-8"))
    except Exception as exc:
        raise RuntimeError("Upload session metadata is corrupted") from exc
    if not part_path.exists():
        raise RuntimeError("Upload session data file is missing")
    return meta, meta_path, part_path


def _save_upload_session(meta: dict[str, Any], meta_path: Path) -> None:
    meta["updated_at"] = datetime.now(timezone.utc).isoformat()
    tmp = meta_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(meta, ensure_ascii=False, separators=(",", ":")), "utf-8")
    os.replace(tmp, meta_path)

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

TEXT_EXTENSIONS = {
    ".txt", ".md", ".markdown", ".csv", ".tsv", ".json", ".jsonl",
    ".xml", ".html", ".htm", ".log", ".ini", ".cfg", ".conf", ".yaml", ".yml"
}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
LEGACY_OFFICE_EXTENSIONS = {".doc", ".xls", ".ppt"}
IMAGE_WRITE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tif", ".tiff"}
IMAGE_EXTENSIONS = IMAGE_WRITE_EXTENSIONS | {".heic", ".heif", ".avif"}
PRO100_EXTENSIONS = {".sto", ".meb", ".shp", ".ese", ".mse"}
BAZIS_EXTENSIONS = {
    ".b3d", ".f3d", ".fr3d", ".bxf", ".bxf2", ".salon", ".k3bz", ".c3d",
    ".step", ".stp", ".sat", ".jt", ".dxf", ".dwg", ".3ds",
    ".wrl", ".wrz", ".obj", ".x3d", ".x3dv", ".md3", ".dae", ".stl"
}
KNOWN_BINARY_EXTENSIONS = {
    ".docx", ".pdf", ".pptx", ".xlsx", ".xlsm", ".xltx", ".xltm",
    ".doc", ".xls", ".ppt", ".zip", ".rar", ".7z",
    ".mp3", ".wav", ".m4a", ".flac", ".mp4", ".mov", ".avi", ".mkv"
} | IMAGE_EXTENSIONS | PRO100_EXTENSIONS | BAZIS_EXTENSIONS


def _file_extension(path: str) -> str:
    return posixpath.splitext(_basename(path))[1].casefold()


def _ensure_text_safe(path: str) -> None:
    ext = _file_extension(path)
    if ext in KNOWN_BINARY_EXTENSIONS:
        raise ValueError(
            f"{ext} is a binary format. Do not use text read/write tools on it; "
            f"use read_file or the format-specific tool, and binary overwrite only when intentionally replacing the whole file."
        )


def _truncate_text(text: str, max_chars: int) -> tuple[str, bool]:
    max_chars = min(max(int(max_chars), 1), 500_000)
    return text[:max_chars], len(text) > max_chars


def _read_docx_bytes(data: bytes, max_chars: int = 100000) -> dict[str, Any]:
    doc = Document(io.BytesIO(data))
    parts: list[str] = []
    paragraph_count = 0
    table_count = len(doc.tables)
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if text:
            paragraph_count += 1
            parts.append(text)
    for t_idx, table in enumerate(doc.tables, start=1):
        parts.append(f"[TABLE {t_idx}]")
        for row in table.rows:
            parts.append(" | ".join(cell.text.strip() for cell in row.cells))
    full = "\n".join(parts)
    text, truncated = _truncate_text(full, max_chars)
    return {
        "format": "docx",
        "paragraphs": paragraph_count,
        "tables": table_count,
        "chars_total": len(full),
        "truncated": truncated,
        "text": text,
    }


def _read_pdf_bytes(data: bytes, max_pages: int = 50, max_chars: int = 100000) -> dict[str, Any]:
    reader = PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        try:
            unlocked = reader.decrypt("")
        except Exception:
            unlocked = 0
        if not unlocked:
            raise ValueError("PDF is encrypted/password-protected and cannot be read without a password")
    max_pages = min(max(int(max_pages), 1), 500)
    page_total = len(reader.pages)
    pages_read = min(page_total, max_pages)
    parts: list[str] = []
    for idx in range(pages_read):
        try:
            page_text = reader.pages[idx].extract_text() or ""
        except Exception as exc:
            page_text = f"[Page {idx + 1}: text extraction failed: {type(exc).__name__}]"
        parts.append(f"[PAGE {idx + 1}]\n{page_text.strip()}")
    full = "\n\n".join(parts)
    text, truncated_chars = _truncate_text(full, max_chars)
    return {
        "format": "pdf",
        "pages_total": page_total,
        "pages_read": pages_read,
        "pages_truncated": pages_read < page_total,
        "chars_total": len(full),
        "truncated": truncated_chars or pages_read < page_total,
        "text": text,
        "note": "Image-only/scanned PDFs may require OCR; this reader extracts embedded text only.",
    }


def _read_pptx_bytes(data: bytes, max_slides: int = 100, max_chars: int = 100000) -> dict[str, Any]:
    prs = Presentation(io.BytesIO(data))
    max_slides = min(max(int(max_slides), 1), 500)
    total = len(prs.slides)
    slides_read = min(total, max_slides)
    parts: list[str] = []
    for idx, slide in enumerate(list(prs.slides)[:slides_read], start=1):
        slide_parts: list[str] = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and str(getattr(shape, "text", "")).strip():
                slide_parts.append(str(shape.text).strip())
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    slide_parts.append(" | ".join(cell.text.strip() for cell in row.cells))
        parts.append(f"[SLIDE {idx}]\n" + "\n".join(slide_parts))
    full = "\n\n".join(parts)
    text, truncated_chars = _truncate_text(full, max_chars)
    return {
        "format": "pptx",
        "slides_total": total,
        "slides_read": slides_read,
        "slides_truncated": slides_read < total,
        "chars_total": len(full),
        "truncated": truncated_chars or slides_read < total,
        "text": text,
    }


def _decode_text_bytes(data: bytes, encoding: str = "utf-8") -> tuple[str, str]:
    candidates = [encoding, "utf-8-sig", "utf-8", "cp1251"]
    seen = set()
    for enc in candidates:
        if enc in seen:
            continue
        seen.add(enc)
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode(encoding, errors="replace"), f"{encoding} (replacement chars used)"


def _read_csv_bytes(data: bytes, max_rows: int = 500, encoding: str = "utf-8") -> dict[str, Any]:
    text, used_encoding = _decode_text_bytes(data, encoding)
    max_rows = min(max(int(max_rows), 1), 5000)
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        delimiter = dialect.delimiter
    except Exception:
        delimiter = ";" if sample.count(";") > sample.count(",") else ","
    rows = []
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    for idx, row in enumerate(reader):
        if idx >= max_rows:
            break
        rows.append(row)
    return {
        "format": "csv",
        "encoding": used_encoding,
        "delimiter": delimiter,
        "rows_returned": len(rows),
        "truncated": len(rows) >= max_rows and text.count("\n") + 1 > max_rows,
        "rows": rows,
    }


def _read_excel_bytes(data: bytes, path: str, sheet: str = "", max_rows: int = 200) -> dict[str, Any]:
    keep_vba = _file_extension(path) in {".xlsm", ".xltm"}
    wb = load_workbook(io.BytesIO(data), keep_vba=keep_vba, read_only=True, data_only=False)
    if sheet and sheet not in wb.sheetnames:
        raise ValueError(f"Sheet not found: {sheet}")
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    max_rows = min(max(int(max_rows), 1), 5000)
    rows = []
    for idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if idx > max_rows:
            break
        rows.append(list(row))
    return {
        "format": "excel",
        "sheets": wb.sheetnames,
        "selected_sheet": ws.title,
        "rows_returned": len(rows),
        "rows": rows,
    }


def _read_text_bytes(data: bytes, max_chars: int = 100000, encoding: str = "utf-8") -> dict[str, Any]:
    full, used_encoding = _decode_text_bytes(data, encoding)
    text, truncated = _truncate_text(full, max_chars)
    return {
        "format": "text",
        "encoding": used_encoding,
        "chars_total": len(full),
        "truncated": truncated,
        "text": text,
    }


def _image_output_format(path: str) -> tuple[str, str]:
    ext = _file_extension(path)
    mapping = {
        ".jpg": ("JPEG", "jpeg"), ".jpeg": ("JPEG", "jpeg"),
        ".png": ("PNG", "png"), ".webp": ("WEBP", "webp"),
        ".bmp": ("BMP", "bmp"), ".gif": ("GIF", "gif"),
        ".tif": ("TIFF", "tiff"), ".tiff": ("TIFF", "tiff"),
    }
    return mapping.get(ext, ("PNG", "png"))


def _prepare_image_bytes(data: bytes, output_path: str, max_dimension: int = 2048, quality: int = 90) -> tuple[bytes, dict[str, Any], str]:
    max_dimension = min(max(int(max_dimension), 128), 4096)
    quality = min(max(int(quality), 20), 100)
    with PILImage.open(io.BytesIO(data)) as src:
        original_format = str(src.format or "unknown")
        original_size = src.size
        img = ImageOps.exif_transpose(src).copy()
        if max(img.size) > max_dimension:
            img.thumbnail((max_dimension, max_dimension), PILImage.Resampling.LANCZOS)
        save_format, mcp_format = _image_output_format(output_path)
        if save_format == "JPEG" and img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        out = io.BytesIO()
        kwargs: dict[str, Any] = {}
        if save_format in {"JPEG", "WEBP"}:
            kwargs["quality"] = quality
        img.save(out, format=save_format, **kwargs)
        meta = {
            "original_format": original_format,
            "original_width": original_size[0],
            "original_height": original_size[1],
            "returned_width": img.size[0],
            "returned_height": img.size[1],
            "returned_format": save_format,
        }
        return out.getvalue(), meta, mcp_format


def _download_range_bytes(path: str, offset: int, length: int) -> tuple[bytes, int | None]:
    normalized = _normalize_path(path)
    offset = max(int(offset), 0)
    length = min(max(int(length), 1), 2_000_000)
    info = _resource(normalized, fields="size")
    total_size = info.get("size")
    r = requests.get(
        f"{YANDEX_API}/resources/download",
        headers=yandex_headers(),
        params={"path": normalized},
        timeout=30,
    )
    _raise(r)
    href = r.json().get("href")
    if not href:
        raise RuntimeError(f"Yandex Disk did not return a download URL for {normalized}")
    end = offset + length - 1
    with requests.get(href, headers={"Range": f"bytes={offset}-{end}"}, timeout=(15, 180), stream=True, allow_redirects=True) as d:
        _raise(d)
        out = io.BytesIO()
        if d.status_code == 206:
            for chunk in d.iter_content(chunk_size=256 * 1024):
                if chunk:
                    out.write(chunk)
                    if out.tell() >= length:
                        break
            return out.getvalue()[:length], int(total_size) if total_size is not None else None
        skipped = 0
        collected = 0
        for chunk in d.iter_content(chunk_size=256 * 1024):
            if not chunk:
                continue
            if skipped + len(chunk) <= offset:
                skipped += len(chunk)
                continue
            start = max(offset - skipped, 0)
            piece = chunk[start:]
            need = length - collected
            out.write(piece[:need])
            collected += min(len(piece), need)
            skipped += len(chunk)
            if collected >= length:
                break
        return out.getvalue(), int(total_size) if total_size is not None else None


@mcp.tool()
def file_capabilities(path: str) -> dict[str, Any]:
    """Report what this bridge can safely do with a file type."""
    ext = _file_extension(path)
    normalized = _normalize_path(path)
    if ext == ".docx":
        return {"path": normalized, "extension": ext, "read": True, "structured_read": "read_docx/read_file", "edit": "binary overwrite only; no safe rich-text editor in v7.6", "server_download": True}
    if ext == ".pdf":
        return {"path": normalized, "extension": ext, "read": True, "structured_read": "read_pdf/read_file", "edit": False, "server_download": True, "note": "Scanned PDFs may need OCR."}
    if ext == ".pptx":
        return {"path": normalized, "extension": ext, "read": True, "structured_read": "read_pptx/read_file", "edit": "binary overwrite only; no safe slide editor in v7.6", "server_download": True}
    if ext in EXCEL_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": True, "structured_read": "read_excel/read_file", "edit": True, "edit_tools": "Excel cell/range/row/sheet tools", "server_download": True}
    if ext in TEXT_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": True, "structured_read": "read_text_file/read_file", "edit": True, "edit_tools": "write_text_file/append_text_file/replace_text", "server_download": True}
    if ext in IMAGE_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": True, "visual_read": "read_image", "metadata": "image_info", "edit": True, "edit_tools": "edit_image_basic or whole-file upload/overwrite", "server_download": True, "note": "read_image returns MCP image content for visual analysis. HEIC/HEIF/AVIF can be read when decoder support is available; save technical edits as JPG/PNG/WEBP/BMP/GIF/TIFF. Client image-content support should be tested after deployment."}
    if ext in PRO100_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": "opaque binary transfer only", "edit": "only by PRO100 or whole-file replacement", "server_download": True, "transfer_tools": "download_file_base64/download_file_chunk_base64/upload_file/overwrite_file/begin_chunked_upload/upload_file_chunk_base64/finish_chunked_upload", "note": "PRO100 native project/library formats are proprietary; this bridge preserves bytes but does not parse project semantics."}
    if ext in BAZIS_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": "opaque binary transfer only", "edit": "only by BAZIS/CAD software or whole-file replacement", "server_download": True, "transfer_tools": "download_file_base64/download_file_chunk_base64/upload_file/overwrite_file/begin_chunked_upload/upload_file_chunk_base64/finish_chunked_upload", "note": "BAZIS/CAD native formats are preserved as binary; parsing/editing their internal model is not attempted by the bridge."}
    if ext in LEGACY_OFFICE_EXTENSIONS:
        return {"path": normalized, "extension": ext, "read": False, "edit": False, "server_download": True, "note": "Legacy binary Office format is not parsed by v7.6; convert to DOCX/XLSX/PPTX."}
    return {"path": normalized, "extension": ext, "read": False, "edit": "binary upload/overwrite only", "server_download": True, "note": "Unknown/binary format. Small files can be returned as base64 with download_file_base64."}


@mcp.tool()
def image_info(path: str) -> dict[str, Any]:
    """Read image metadata safely without changing the file."""
    normalized = _normalize_path(path)
    if _file_extension(normalized) not in IMAGE_EXTENSIONS:
        raise ValueError("image_info supports common image files only")
    data = _download_bytes(normalized, max_bytes=50 * 1024 * 1024)
    with PILImage.open(io.BytesIO(data)) as img:
        return {
            "ok": True,
            "path": normalized,
            "format": img.format,
            "mode": img.mode,
            "width": img.size[0],
            "height": img.size[1],
            "frames": int(getattr(img, "n_frames", 1)),
            "bytes": len(data),
        }


@mcp.tool()
def read_image(path: str, max_dimension: int = 2048, quality: int = 90) -> MCPImage:
    """Return an image from Yandex Disk as MCP image content for model vision.

    The bridge downloads the image on Render and returns image bytes through MCP,
    avoiding downloader.disk.yandex.ru in the ChatGPT runtime. Large images are
    downscaled for reliable analysis; the original file on Disk is untouched.
    """
    normalized = _normalize_path(path)
    if _file_extension(normalized) not in IMAGE_EXTENSIONS:
        raise ValueError("read_image supports JPG/JPEG/PNG/WEBP/BMP/GIF/TIFF images only")
    data = _download_bytes(normalized, max_bytes=50 * 1024 * 1024)
    prepared, _meta, mcp_format = _prepare_image_bytes(data, normalized, max_dimension=max_dimension, quality=quality)
    return MCPImage(data=prepared, format=mcp_format)


@mcp.tool()
def edit_image_basic(
    path: str,
    destination_path: str,
    rotation_degrees: int = 0,
    resize_width: int = 0,
    resize_height: int = 0,
    flip_horizontal: bool = False,
    flip_vertical: bool = False,
    crop_left: int = -1,
    crop_top: int = -1,
    crop_right: int = -1,
    crop_bottom: int = -1,
    quality: int = 90,
    overwrite: bool = False,
    make_backup: bool = True,
) -> dict[str, Any]:
    """Perform deterministic image edits on Render and save back to Yandex Disk.

    Supports crop, resize, rotation, flips and format conversion based on the
    destination extension. This is for technical edits, not generative/object-level edits.
    """
    source = _normalize_path(path)
    dest = _normalize_path(destination_path)
    if _file_extension(source) not in IMAGE_EXTENSIONS or _file_extension(dest) not in IMAGE_WRITE_EXTENSIONS:
        raise ValueError("Source must be a supported image; destination must be JPG/JPEG/PNG/WEBP/BMP/GIF/TIFF")
    data = _download_bytes(source, max_bytes=50 * 1024 * 1024)
    quality = min(max(int(quality), 20), 100)
    with PILImage.open(io.BytesIO(data)) as src_img:
        img = ImageOps.exif_transpose(src_img).copy()
        before = img.size
        crop_vals = (crop_left, crop_top, crop_right, crop_bottom)
        if any(v >= 0 for v in crop_vals):
            if not all(v >= 0 for v in crop_vals):
                raise ValueError("For cropping, provide all four crop coordinates")
            if not (0 <= crop_left < crop_right <= img.width and 0 <= crop_top < crop_bottom <= img.height):
                raise ValueError("Invalid crop rectangle")
            img = img.crop((crop_left, crop_top, crop_right, crop_bottom))
        if flip_horizontal:
            img = ImageOps.mirror(img)
        if flip_vertical:
            img = ImageOps.flip(img)
        if rotation_degrees:
            img = img.rotate(-int(rotation_degrees), expand=True)
        if resize_width > 0 or resize_height > 0:
            if resize_width <= 0:
                resize_width = max(1, round(img.width * (resize_height / img.height)))
            if resize_height <= 0:
                resize_height = max(1, round(img.height * (resize_width / img.width)))
            resize_width = min(max(int(resize_width), 1), 12000)
            resize_height = min(max(int(resize_height), 1), 12000)
            img = img.resize((resize_width, resize_height), PILImage.Resampling.LANCZOS)
        save_format, _mcp_format = _image_output_format(dest)
        if save_format == "JPEG" and img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        out = io.BytesIO()
        kwargs: dict[str, Any] = {}
        if save_format in {"JPEG", "WEBP"}:
            kwargs["quality"] = quality
        img.save(out, format=save_format, **kwargs)
        final_bytes = out.getvalue()
        after = img.size
    backup = None
    if overwrite and _resource_exists(dest) and make_backup:
        backup = backup_file(dest).get("backup_path")
    result = _upload_bytes(dest, final_bytes, overwrite=overwrite)
    result.update({
        "source_path": source,
        "destination_path": dest,
        "before_size": list(before),
        "after_size": list(after),
        "format": save_format,
        "backup_path": backup,
        "note": "Basic deterministic image edit completed. Generative edits require an image model outside this bridge.",
    })
    return result


@mcp.tool()
def read_docx(path: str, max_chars: int = 100000) -> dict[str, Any]:
    """Read DOCX on the Render server and return extracted paragraphs/tables."""
    normalized = _normalize_path(path)
    if _file_extension(normalized) != ".docx":
        raise ValueError("read_docx supports .docx files only")
    result = _read_docx_bytes(_download_bytes(normalized), max_chars=max_chars)
    result["path"] = normalized
    return result


@mcp.tool()
def read_pdf(path: str, max_pages: int = 50, max_chars: int = 100000) -> dict[str, Any]:
    """Read embedded text from a PDF on the Render server."""
    normalized = _normalize_path(path)
    if _file_extension(normalized) != ".pdf":
        raise ValueError("read_pdf supports .pdf files only")
    result = _read_pdf_bytes(_download_bytes(normalized), max_pages=max_pages, max_chars=max_chars)
    result["path"] = normalized
    return result


@mcp.tool()
def read_pptx(path: str, max_slides: int = 100, max_chars: int = 100000) -> dict[str, Any]:
    """Read text and tables from PPTX on the Render server."""
    normalized = _normalize_path(path)
    if _file_extension(normalized) != ".pptx":
        raise ValueError("read_pptx supports .pptx files only")
    result = _read_pptx_bytes(_download_bytes(normalized), max_slides=max_slides, max_chars=max_chars)
    result["path"] = normalized
    return result


@mcp.tool()
def read_csv_file(path: str, max_rows: int = 500, encoding: str = "utf-8") -> dict[str, Any]:
    normalized = _normalize_path(path)
    if _file_extension(normalized) not in {".csv", ".tsv"}:
        raise ValueError("read_csv_file supports .csv/.tsv files only")
    result = _read_csv_bytes(_download_bytes(normalized), max_rows=max_rows, encoding=encoding)
    result["path"] = normalized
    return result


@mcp.tool()
def read_file(path: str, max_chars: int = 100000, max_pages: int = 50, max_rows: int = 500, sheet: str = "", encoding: str = "utf-8") -> dict[str, Any]:
    """Unified server-side reader for common business file formats.

    Supports DOCX, PDF, PPTX, XLSX/XLSM/XLTX/XLTM, CSV/TSV and common text
    formats. The actual file is downloaded by Render, so ChatGPT does not need
    to follow Yandex's temporary downloader URL.
    """
    normalized = _normalize_path(path)
    ext = _file_extension(normalized)
    data = _download_bytes(normalized)
    if ext == ".docx":
        result = _read_docx_bytes(data, max_chars=max_chars)
    elif ext == ".pdf":
        result = _read_pdf_bytes(data, max_pages=max_pages, max_chars=max_chars)
    elif ext == ".pptx":
        result = _read_pptx_bytes(data, max_slides=max_pages, max_chars=max_chars)
    elif ext in EXCEL_EXTENSIONS:
        result = _read_excel_bytes(data, normalized, sheet=sheet, max_rows=max_rows)
    elif ext in {".csv", ".tsv"}:
        result = _read_csv_bytes(data, max_rows=max_rows, encoding=encoding)
    elif ext in TEXT_EXTENSIONS:
        result = _read_text_bytes(data, max_chars=max_chars, encoding=encoding)
    elif ext in LEGACY_OFFICE_EXTENSIONS:
        raise ValueError(f"Legacy Office format {ext} is not supported for parsing; convert it to a modern Office format")
    else:
        raise ValueError(f"Unsupported file format {ext or '(no extension)'}; use file_capabilities for details")
    result["path"] = normalized
    result["extension"] = ext
    return result


@mcp.tool()
def download_file_base64(path: str, max_bytes: int = 1_000_000) -> dict[str, Any]:
    """Return a small binary file directly through MCP as base64.

    This avoids external temporary URLs, but is intentionally capped because
    large base64 payloads are inefficient in ChatGPT tool responses. Use
    read_file/read_docx/read_pdf/read_pptx for analysis of larger documents.
    """
    normalized = _normalize_path(path)
    max_bytes = min(max(int(max_bytes), 1), 4_000_000)
    data = _download_bytes(normalized, max_bytes=max_bytes)
    return {
        "path": normalized,
        "bytes": len(data),
        "base64": base64.b64encode(data).decode("ascii"),
        "warning": "For analysis prefer read_file; base64 is intended for small binary transfers only.",
    }


@mcp.tool()
def download_file_chunk_base64(path: str, offset: int = 0, length: int = 1_000_000) -> dict[str, Any]:
    """Return one binary chunk through MCP for arbitrary/proprietary files.

    Useful for PRO100, BAZIS and other formats that cannot be parsed by the bridge.
    Call repeatedly with next_offset until has_more is false. Each chunk is capped
    at 2 MB to keep MCP responses manageable.
    """
    normalized = _normalize_path(path)
    offset = max(int(offset), 0)
    length = min(max(int(length), 1), 2_000_000)
    data, total_size = _download_range_bytes(normalized, offset, length)
    next_offset = offset + len(data)
    has_more = total_size is None or next_offset < total_size
    return {
        "path": normalized,
        "offset": offset,
        "bytes_returned": len(data),
        "total_size": total_size,
        "next_offset": next_offset if has_more else None,
        "has_more": has_more,
        "base64": base64.b64encode(data).decode("ascii"),
        "note": "Opaque binary transfer. Do not interpret proprietary CAD/project bytes as text.",
    }


@mcp.tool()
def get_file_download_url(path: str) -> dict[str, str]:
    path = _normalize_path(path)
    r = requests.get(f"{YANDEX_API}/resources/download", headers=yandex_headers(), params={"path": path}, timeout=30)
    _raise(r)
    return {
        "path": path,
        "download_url": r.json()["href"],
        "note": "Temporary Yandex URL for external clients. For ChatGPT analysis prefer read_file/read_docx/read_pdf/read_pptx; some runtimes cannot fetch downloader.disk.yandex.ru directly.",
    }

@mcp.tool()
def read_text_file(path: str, max_chars: int = 50000, encoding: str = "utf-8") -> dict[str, Any]:
    path = _normalize_path(path)
    _ensure_text_safe(path)
    result = _read_text_bytes(_download_bytes(path), max_chars=max_chars, encoding=encoding)
    result["path"] = path
    return result

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
        "server": "Yandex Disk — Я Мебель v7.6.3a strict-base64-2k fix",
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


def _mcp_registry_diagnostics() -> dict[str, Any]:
    """Return a safe snapshot of FastMCP's in-process tool registry.

    This intentionally avoids network calls and secrets. It is designed to
    diagnose schema/cache mismatches by reporting what the running Python
    process has actually registered before the MCP transport publishes it.
    """
    expected_chunked = [
        "begin_chunked_upload",
        "upload_file_chunk_base64",
        "finish_chunked_upload",
        "abort_chunked_upload",
    ]
    result: dict[str, Any] = {
        "registry_found": False,
        "registered_tool_count": None,
        "registered_tool_names": [],
        "chunked_upload_expected": expected_chunked,
        "chunked_upload_registered": {},
        "probe": {},
    }
    try:
        manager = getattr(mcp, "_tool_manager", None)
        result["probe"]["mcp_has_tool_manager"] = manager is not None
        if manager is None:
            result["probe"]["mcp_attributes"] = sorted(
                name for name in vars(mcp).keys() if "tool" in name.casefold()
            )
            return result

        result["probe"]["tool_manager_type"] = type(manager).__name__
        tools = getattr(manager, "_tools", None)
        result["probe"]["tool_manager_has__tools"] = tools is not None

        names: list[str] = []
        if isinstance(tools, dict):
            names = [str(name) for name in tools.keys()]
            result["probe"]["registry_container_type"] = "dict"
        elif tools is not None:
            result["probe"]["registry_container_type"] = type(tools).__name__
            try:
                for item in tools:
                    name = getattr(item, "name", None)
                    if name:
                        names.append(str(name))
            except TypeError:
                pass

        # Fallback: inspect only manager attributes whose values are obvious
        # tool containers. This is deliberately shallow to avoid serializing
        # internal objects or user data.
        if not names:
            candidates: dict[str, Any] = {}
            for attr_name, value in vars(manager).items():
                if "tool" not in attr_name.casefold():
                    continue
                if isinstance(value, dict):
                    candidate_names = [str(k) for k in value.keys()]
                    candidates[attr_name] = {
                        "type": "dict",
                        "count": len(candidate_names),
                    }
                    if len(candidate_names) > len(names):
                        names = candidate_names
                elif isinstance(value, (list, tuple, set)):
                    candidate_names = []
                    for item in value:
                        name = getattr(item, "name", None)
                        if name:
                            candidate_names.append(str(name))
                    candidates[attr_name] = {
                        "type": type(value).__name__,
                        "count": len(value),
                    }
                    if len(candidate_names) > len(names):
                        names = candidate_names
            result["probe"]["fallback_tool_containers"] = candidates

        names = sorted(set(names))
        result["registry_found"] = bool(names)
        result["registered_tool_count"] = len(names) if names else None
        result["registered_tool_names"] = names
        result["chunked_upload_registered"] = {
            name: name in names for name in expected_chunked
        }
        result["all_chunked_upload_registered"] = bool(names) and all(
            name in names for name in expected_chunked
        )
        return result
    except Exception as exc:
        result["diagnostic_error_type"] = type(exc).__name__
        result["diagnostic_error"] = str(exc)[:500]
        return result


@mcp.tool()
def bridge_version() -> dict[str, Any]:
    return {
        "ok": True,
        "bridge_version": APP_VERSION,
        "mcp_registry": _mcp_registry_diagnostics(),
    }


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
def begin_chunked_upload(
    path: str,
    total_size: int,
    overwrite: bool = False,
    make_backup: bool = True,
    sha256: str = "",
) -> dict[str, Any]:
    """Start a resumable MCP upload for a generated or large binary file.

    After this call, send base64 chunks in strict byte order with
    upload_file_chunk_base64, then call finish_chunked_upload. The original
    Yandex file is not modified until finish succeeds. Staging is ephemeral and
    an in-progress session can be lost if Render restarts.
    """
    _cleanup_stale_uploads()
    normalized = _normalize_path(path)
    total_size = int(total_size)
    if total_size < 0:
        raise ValueError("total_size must be >= 0")
    if total_size > MAX_CHUNK_UPLOAD_BYTES:
        raise ValueError(f"File is too large for chunked staging: {total_size} bytes; limit is {MAX_CHUNK_UPLOAD_BYTES}")
    expected_sha = sha256.strip().casefold()
    if expected_sha and (len(expected_sha) != 64 or any(c not in "0123456789abcdef" for c in expected_sha)):
        raise ValueError("sha256 must be a 64-character hexadecimal digest")
    exists = _resource_exists(normalized)
    if exists and not overwrite:
        raise FileExistsError(f"Destination already exists: {normalized}; set overwrite=true to replace it")
    upload_id = str(uuid.uuid4())
    meta_path, part_path = _upload_session_paths(upload_id)
    part_path.write_bytes(b"")
    meta = {
        "upload_id": upload_id,
        "path": normalized,
        "total_size": total_size,
        "bytes_received": 0,
        "overwrite": bool(overwrite),
        "make_backup": bool(make_backup),
        "expected_sha256": expected_sha or None,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }
    _save_upload_session(meta, meta_path)
    return {
        "ok": True,
        **meta,
        "max_chunk_bytes": min(max(MAX_CHUNK_BYTES, 1), SAFE_MCP_CHUNK_CAP_BYTES),
        "recommended_chunk_bytes": RECOMMENDED_MCP_CHUNK_BYTES,
        "transport_profile": "strict-base64-fixed-2k-v7.6.3",
        "next_offset": 0,
        "note": "Send raw file bytes as base64 chunks in strict order. Do not base64-encode an already base64 string.",
    }


def _decode_base64_transport_chunk(content_base64: str, current: int) -> tuple[bytes, dict[str, Any]]:
    """Decode exactly one canonical Base64 chunk.

    No whitespace trimming, URL-safe normalization, padding repair, data-URI
    handling, or repeated decoding is performed. Trailing '=' padding is
    preserved exactly as received and validated by Python's strict decoder.
    """
    if not isinstance(content_base64, str):
        raise TypeError("content_base64 must be a string")
    if not content_base64:
        raise ValueError(
            f"Invalid base64 chunk: empty payload. Session was preserved; retry offset {current}."
        )
    try:
        chunk = base64.b64decode(content_base64, validate=True)
    except Exception as exc:
        raise ValueError(
            f"Invalid canonical base64 chunk: {exc}. Session was preserved; retry offset {current}. "
            "Send only the exact standard Base64 string, including any trailing '=' padding."
        ) from exc
    return chunk, {
        "input_chars": len(content_base64),
        "strict_decode": True,
        "padding_chars": len(content_base64) - len(content_base64.rstrip("=")),
        "normalization_applied": False,
        "decode_passes": 1,
    }


@mcp.tool()
def upload_file_chunk_base64(upload_id: str, offset: int, content_base64: str) -> dict[str, Any]:
    """Append one strict standard-Base64 chunk to an active upload session.

    Chunks are fixed at RECOMMENDED_MCP_CHUNK_BYTES (2048 bytes) except the
    final chunk, which must decode to exactly the remaining file size. This
    makes truncation or partial decoding fail before any bytes are appended.
    """
    meta, meta_path, part_path = _load_upload_session(upload_id)
    offset = int(offset)
    current = int(meta.get("bytes_received", 0))
    if offset != current:
        raise ValueError(f"Wrong chunk offset: expected {current}, got {offset}")

    total_size = int(meta["total_size"])
    remaining_before = total_size - current
    if remaining_before <= 0:
        raise ValueError(f"Upload already has all declared bytes: {total_size}")

    expected_chunk_bytes = min(RECOMMENDED_MCP_CHUNK_BYTES, remaining_before)
    chunk, transport_info = _decode_base64_transport_chunk(content_base64, current)

    if len(chunk) != expected_chunk_bytes:
        raise ValueError(
            f"Decoded chunk size mismatch at offset {current}: "
            f"expected {expected_chunk_bytes} bytes, got {len(chunk)}. "
            "Session was preserved and no bytes from this chunk were appended."
        )

    max_chunk = min(max(MAX_CHUNK_BYTES, 1), SAFE_MCP_CHUNK_CAP_BYTES)
    if len(chunk) > max_chunk:
        raise ValueError(f"Chunk is too large: {len(chunk)} bytes; maximum is {max_chunk}")

    with part_path.open("ab") as fh:
        fh.write(chunk)
        fh.flush()
        os.fsync(fh.fileno())

    current += len(chunk)
    meta["bytes_received"] = current
    _save_upload_session(meta, meta_path)
    complete = current == total_size
    return {
        "ok": True,
        "upload_id": upload_id,
        "path": meta["path"],
        "bytes_received": current,
        "total_size": total_size,
        "next_offset": None if complete else current,
        "complete": complete,
        "remaining_bytes": total_size - current,
        "expected_chunk_bytes": expected_chunk_bytes,
        "accepted_chunk_bytes": len(chunk),
        "chunk_sha256": hashlib.sha256(chunk).hexdigest(),
        "base64_transport": transport_info,
        "recommended_chunk_bytes": RECOMMENDED_MCP_CHUNK_BYTES,
    }


@mcp.tool()
def finish_chunked_upload(upload_id: str) -> dict[str, Any]:
    """Validate and atomically finish a staged chunked upload to Yandex Disk."""
    meta, meta_path, part_path = _load_upload_session(upload_id)
    total_size = int(meta["total_size"])
    actual_size = part_path.stat().st_size
    recorded_size = int(meta.get("bytes_received", 0))
    if actual_size != recorded_size or actual_size != total_size:
        raise ValueError(
            f"Upload is incomplete: actual={actual_size}, recorded={recorded_size}, expected={total_size}"
        )
    digest = hashlib.sha256()
    with part_path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    actual_sha = digest.hexdigest()
    expected_sha = meta.get("expected_sha256")
    if expected_sha and actual_sha != expected_sha:
        raise ValueError(f"SHA-256 mismatch: expected {expected_sha}, got {actual_sha}")
    path = _normalize_path(meta["path"])
    overwrite = bool(meta.get("overwrite", False))
    make_backup = bool(meta.get("make_backup", True))
    # Re-check destination at commit time to prevent accidental replacement if
    # another client created the file while chunks were being staged.
    exists_now = _resource_exists(path)
    if exists_now and not overwrite:
        raise FileExistsError(f"Destination appeared during upload: {path}; session was not committed")
    backup = _backup_if_requested(path, make_backup) if overwrite else None
    try:
        result = _upload_local_file(path, part_path, overwrite=overwrite)
    except Exception:
        # Keep the staged data so finish can be retried after a transient error.
        raise
    result.update({
        "upload_id": upload_id,
        "sha256": actual_sha,
        "backup_path": backup,
        "chunked": True,
    })
    meta_path.unlink(missing_ok=True)
    part_path.unlink(missing_ok=True)
    return result


@mcp.tool()
def abort_chunked_upload(upload_id: str) -> dict[str, Any]:
    """Discard an unfinished chunked upload without touching Yandex Disk."""
    meta, meta_path, part_path = _load_upload_session(upload_id)
    bytes_received = part_path.stat().st_size if part_path.exists() else int(meta.get("bytes_received", 0))
    meta_path.unlink(missing_ok=True)
    part_path.unlink(missing_ok=True)
    return {
        "ok": True,
        "upload_id": upload_id,
        "path": meta.get("path"),
        "discarded_bytes": bytes_received,
        "yandex_file_modified": False,
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
    _ensure_text_safe(path)
    data = text.encode(encoding)
    backup = _backup_if_requested(path, make_backup) if overwrite else None
    result = _upload_bytes(path, data, overwrite)
    result["backup_path"] = backup
    return result

@mcp.tool()
def append_text_file(path: str, text: str, encoding: str = "utf-8", make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    _ensure_text_safe(path)
    existing = _download_bytes(path).decode(encoding, errors="replace")
    data = (existing + text).encode(encoding)
    backup = _backup_if_requested(path, make_backup)
    result = _upload_bytes(path, data, True)
    result["backup_path"] = backup
    return result

@mcp.tool()
def replace_text(path: str, old_text: str, new_text: str, replace_all: bool = True, encoding: str = "utf-8", make_backup: bool = True) -> dict[str, Any]:
    path = _normalize_path(path)
    _ensure_text_safe(path)
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
