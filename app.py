"""
bitte レシピ写真管理アプリ
FastAPI + ローカル実行
"""
import json
import os
import re
import sys
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "recipes_all.xlsx"
PHOTOS_DIR = BASE_DIR / "recipe_photos"
PHOTOS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="bitte Recipe Photo Manager")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _slug(name: str) -> str:
    """レシピ名をディレクトリ名に変換（パス区切り文字を除去）。"""
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def _get_recipes() -> list[str]:
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


@app.get("/api/recipes")
def list_recipes():
    recipes = _get_recipes()
    result = []
    for name in recipes:
        slug = _slug(name)
        d = PHOTOS_DIR / slug
        photos = sorted([f.name for f in d.iterdir() if f.is_file()]) if d.exists() else []
        result.append({"name": name, "slug": slug, "photo_count": len(photos)})
    return result


@app.get("/api/photos/{slug}")
def list_photos(slug: str):
    d = PHOTOS_DIR / slug
    if not d.exists():
        return []
    return sorted([f.name for f in d.iterdir() if f.is_file()])


@app.post("/api/photos/{slug}")
async def upload_photo(slug: str, file: UploadFile = File(...)):
    # スラグが有効か確認
    recipes = _get_recipes()
    slugs = {_slug(r): r for r in recipes}
    if slug not in slugs:
        raise HTTPException(status_code=404, detail="レシピが見つかりません")

    # 拡張子チェック
    ext = Path(file.filename).suffix.lower()
    if ext not in {".jpg", ".jpeg", ".png", ".webp", ".heic"}:
        raise HTTPException(status_code=400, detail="jpg/png/webp/heic のみ対応しています")

    d = PHOTOS_DIR / slug
    d.mkdir(exist_ok=True)

    # ファイル名重複回避
    dest = d / file.filename
    stem = Path(file.filename).stem
    counter = 1
    while dest.exists():
        dest = d / f"{stem}_{counter}{ext}"
        counter += 1

    content = await file.read()
    dest.write_bytes(content)
    return {"filename": dest.name, "slug": slug}


@app.delete("/api/photos/{slug}/{filename}")
def delete_photo(slug: str, filename: str):
    path = PHOTOS_DIR / slug / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="ファイルが見つかりません")
    path.unlink()
    return {"result": "ok"}


@app.get("/photos/{slug}/{filename}")
def serve_photo(slug: str, filename: str):
    path = PHOTOS_DIR / slug / filename
    if not path.exists():
        raise HTTPException(status_code=404)
    return FileResponse(path)


# --- フロントエンド配信 ---
PUBLIC_DIR = BASE_DIR / "public"
PUBLIC_DIR.mkdir(exist_ok=True)


@app.get("/")
def index():
    return FileResponse(PUBLIC_DIR / "index.html")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8001, reload=True)
