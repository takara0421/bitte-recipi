"""
bitte レシピ写真管理 API - Vercel Serverless
FastAPI + Vercel Blob（写真ストレージ）
"""
import os
import re
from pathlib import Path

import requests as _req
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware

BASE_DIR = Path(__file__).parent.parent
EXCEL_PATH = BASE_DIR / "recipes_all.xlsx"
BLOB_API = "https://blob.vercel-storage.com"

app = FastAPI(title="bitte Recipe Photo Manager")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _token() -> str:
    t = os.environ.get("BLOB_READ_WRITE_TOKEN")
    if not t:
        raise HTTPException(status_code=503, detail="BLOB_READ_WRITE_TOKEN が未設定です")
    return t


def _slug(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def _get_recipes() -> list:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH)
        ws = wb.active
        seen, result = set(), []
        for row in ws.iter_rows(min_row=2, values_only=True):
            name = row[0]
            if name and isinstance(name, str) and name not in seen and not name.startswith("※"):
                seen.add(name)
                result.append(name)
        return sorted(result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel読み込みエラー: {e}")


def _list_blobs(prefix: str) -> list:
    r = _req.get(
        BLOB_API,
        headers={"Authorization": f"Bearer {_token()}"},
        params={"prefix": prefix, "limit": 1000},
        timeout=15,
    )
    r.raise_for_status()
    return r.json().get("blobs", [])


@app.get("/api/recipes")
def list_recipes():
    recipes = _get_recipes()
    all_blobs = _list_blobs("recipes/")
    # slug ごとにカウント
    counts: dict = {}
    for b in all_blobs:
        parts = b.get("pathname", "").split("/")
        if len(parts) >= 3:  # recipes/{slug}/{filename}
            counts[parts[1]] = counts.get(parts[1], 0) + 1
    return [
        {"name": r, "slug": _slug(r), "photo_count": counts.get(_slug(r), 0)}
        for r in recipes
    ]


@app.get("/api/photos/{slug}")
def list_photos(slug: str):
    blobs = _list_blobs(f"recipes/{slug}/")
    return [
        {"url": b["url"], "filename": b["pathname"].split("/")[-1]}
        for b in blobs
    ]


@app.post("/api/photos/{slug}")
async def upload_photo(slug: str, file: UploadFile = File(...)):
    slugs = {_slug(r): r for r in _get_recipes()}
    if slug not in slugs:
        raise HTTPException(status_code=404, detail="レシピが見つかりません")

    ext = Path(file.filename).suffix.lower()
    if ext not in {".jpg", ".jpeg", ".png", ".webp", ".heic"}:
        raise HTTPException(status_code=400, detail="jpg/png/webp/heic のみ対応しています")

    content_type = {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".webp": "image/webp",
        ".heic": "image/heic",
    }.get(ext, "image/jpeg")

    content = await file.read()
    pathname = f"recipes/{slug}/{file.filename}"

    r = _req.put(
        f"{BLOB_API}/{pathname}",
        headers={
            "Authorization": f"Bearer {_token()}",
            "x-api-version": "7",
            "content-type": content_type,
            "x-add-random-suffix": "1",
        },
        data=content,
        timeout=60,
    )
    if not r.ok:
        raise HTTPException(status_code=r.status_code, detail=f"Blob upload error: {r.text}")
    data = r.json()
    return {"url": data["url"], "filename": data["pathname"].split("/")[-1]}


@app.delete("/api/photos")
async def delete_photo(body: dict):
    url = body.get("url")
    if not url:
        raise HTTPException(status_code=400, detail="url required")
    r = _req.delete(
        BLOB_API,
        headers={"Authorization": f"Bearer {_token()}", "content-type": "application/json"},
        json={"urls": [url]},
        timeout=15,
    )
    if not r.ok:
        raise HTTPException(status_code=r.status_code, detail=f"Blob delete error: {r.text}")
    return {"result": "ok"}
