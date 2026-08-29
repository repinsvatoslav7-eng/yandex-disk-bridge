import os
from urllib.parse import quote

import requests
from flask import Flask, jsonify, Response, request

app = Flask(__name__)

YANDEX_TOKEN = os.environ.get("YANDEX_DISK_TOKEN")
API = "https://cloud-api.yandex.net/v1/disk"


def headers():
    return {"Authorization": f"OAuth {YANDEX_TOKEN}"}


@app.get("/")
def home():
    return jsonify({
        "status": "ok",
        "service": "Yandex Disk Bridge"
    })


@app.get("/list")
def list_folder():
    path = request.args.get("path", "/")

    r = requests.get(
        f"{API}/resources",
        headers=headers(),
        params={
            "path": path,
            "limit": 1000,
            "fields": "_embedded.items.name,_embedded.items.type,_embedded.items.path"
        },
        timeout=30
    )

    if not r.ok:
        return jsonify({"error": r.text}), r.status_code

    data = r.json()
    return jsonify(data.get("_embedded", {}).get("items", []))


@app.get("/download")
def download_file():
    path = request.args.get("path")

    if not path:
        return jsonify({"error": "path is required"}), 400

    r = requests.get(
        f"{API}/resources/download",
        headers=headers(),
        params={"path": path},
        timeout=30
    )

    if not r.ok:
        return jsonify({"error": r.text}), r.status_code

    href = r.json().get("href")

    if not href:
        return jsonify({"error": "download link not received"}), 500

    file_response = requests.get(href, stream=True, timeout=60)

    filename = path.split("/")[-1]

    return Response(
        file_response.iter_content(chunk_size=8192),
        content_type=file_response.headers.get(
            "Content-Type",
            "application/octet-stream"
        ),
        headers={
            "Content-Disposition":
                f"attachment; filename*=UTF-8''{quote(filename)}"
        }
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
