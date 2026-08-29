import io
import os
from typing import Any

import requests
from openpyxl import load_workbook
from mcp.server.fastmcp import FastMCP


YANDEX_TOKEN = os.environ["YANDEX_DISK_TOKEN"]
PORT = int(os.environ.get("PORT", "10000"))
MCP_SECRET = os.environ.get("MCP_SECRET", "change-me")

YANDEX_API = "https://cloud-api.yandex.net/v1/disk"

mcp = FastMCP(
    "Yandex Disk",
    instructions=(
        "Read-only access to the user's Yandex Disk. "
        "Use list_folder to browse folders and read_excel to inspect Excel files."
    ),
    host="0.0.0.0",
    port=PORT,
    stateless_http=True,
    json_response=True,
    streamable_http_path=f"/mcp/{MCP_SECRET}",
)


def yandex_headers() -> dict[str, str]:
    return {"Authorization": f"OAuth {YANDEX_TOKEN}"}


def get_download_link(path: str) -> str:
    response = requests.get(
        f"{YANDEX_API}/resources/download",
        headers=yandex_headers(),
        params={"path": path},
        timeout=30,
    )
    response.raise_for_status()

    href = response.json().get("href")
    if not href:
        raise RuntimeError("Yandex Disk did not return a download link.")

    return href


@mcp.tool()
def list_folder(path: str = "/") -> list[dict[str, Any]]:
    """
    List files and folders in a Yandex Disk folder.

    Use paths such as:
    /
    /Филиал Ростов
    /Филиал Ростов/Анализ конкурентов
    """
    response = requests.get(
        f"{YANDEX_API}/resources",
        headers=yandex_headers(),
        params={
            "path": path,
            "limit": 1000,
            "fields": (
                "_embedded.items.name,"
                "_embedded.items.type,"
                "_embedded.items.path,"
                "_embedded.items.size,"
                "_embedded.items.mime_type"
            ),
        },
        timeout=30,
    )
    response.raise_for_status()

    items = response.json().get("_embedded", {}).get("items", [])

    return [
        {
            "name": item.get("name"),
            "type": item.get("type"),
            "path": item.get("path"),
            "size": item.get("size"),
            "mime_type": item.get("mime_type"),
        }
        for item in items
    ]


@mcp.tool()
def get_file_download_url(path: str) -> dict[str, str]:
    """
    Get a temporary download URL for a file on Yandex Disk.
    """
    return {
        "path": path,
        "download_url": get_download_link(path),
    }


@mcp.tool()
def read_excel(
    path: str,
    sheet: str = "",
    max_rows: int = 200,
) -> dict[str, Any]:
    """
    Read an XLSX file from Yandex Disk and return spreadsheet data.

    If sheet is empty, returns the first worksheet.
    max_rows limits the number of returned rows.
    """
    max_rows = max(1, min(max_rows, 500))

    download_url = get_download_link(path)

    response = requests.get(download_url, timeout=90)
    response.raise_for_status()

    workbook = load_workbook(
        io.BytesIO(response.content),
        read_only=True,
        data_only=True,
    )

    sheet_names = workbook.sheetnames

    if not sheet_names:
        return {
            "path": path,
            "sheets": [],
            "rows": [],
        }

    selected_sheet = sheet if sheet in sheet_names else sheet_names[0]
    worksheet = workbook[selected_sheet]

    rows = []

    for index, row in enumerate(
        worksheet.iter_rows(values_only=True),
        start=1,
    ):
        if index > max_rows:
            break

        rows.append([
            value.isoformat() if hasattr(value, "isoformat") else value
            for value in row
        ])

    return {
        "path": path,
        "sheets": sheet_names,
        "selected_sheet": selected_sheet,
        "rows_returned": len(rows),
        "rows": rows,
    }


@mcp.tool()
def read_text_file(
    path: str,
    max_chars: int = 50000,
) -> dict[str, Any]:
    """
    Read a text-like file from Yandex Disk.
    Suitable for txt, md, csv and similar text files.
    """
    max_chars = max(1000, min(max_chars, 200000))

    download_url = get_download_link(path)

    response = requests.get(download_url, timeout=90)
    response.raise_for_status()

    response.encoding = response.apparent_encoding or "utf-8"
    text = response.text[:max_chars]

    return {
        "path": path,
        "characters_returned": len(text),
        "text": text,
    }


if __name__ == "__main__":
    mcp.run(transport="streamable-http")
