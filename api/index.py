"""
bitte レシピ写真管理 API - Vercel Serverless
FastAPI + Google Drive（写真ストレージ）
必要な環境変数: GOOGLE_CREDENTIALS のみ
"""
import io
import json
import os
import re
from pathlib import Path

from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response

BASE_DIR = Path(__file__).parent.parent
EXCEL_PATH = BASE_DIR / "recipes_all.xlsx"
ROOT_FOLDER_NAME = "bitte-recipe-photos"

app = FastAPI(title="bitte Recipe Photo Manager")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== Google Drive 接続 =====

_svc_cache = None
_root_id_cache = None


def _drive():
    global _svc_cache
    if _svc_cache is not None:
        return _svc_cache
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if not creds_json:
        raise HTTPException(status_code=503, detail="GOOGLE_CREDENTIALS が未設定です")
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        creds = Credentials.from_service_account_info(
            json.loads(creds_json),
            scopes=["https://www.googleapis.com/auth/drive"],
        )
        _svc_cache = build("drive", "v3", credentials=creds, cache_discovery=False)
        return _svc_cache
    except Exception as e:
        raise HTTPException(status_code=503, detail=f"Drive接続エラー: {e}")


def _root_id() -> str:
    """ルートフォルダ "bitte-recipe-photos" のIDを取得（なければ作成）。"""
    global _root_id_cache
    if _root_id_cache:
        return _root_id_cache

    # 環境変数で上書き可能
    env_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID")
    if env_id:
        _root_id_cache = env_id
        return _root_id_cache

    svc = _drive()
    q = (
        f"name='{ROOT_FOLDER_NAME}'"
        " and mimeType='application/vnd.google-apps.folder'"
        " and trashed=false"
    )
    res = svc.files().list(q=q, fields="files(id)", pageSize=1).execute()
    files = res.get("files", [])
    if files:
        _root_id_cache = files[0]["id"]
        return _root_id_cache

    # 新規作成してオーナーと共有
    meta = {
        "name": ROOT_FOLDER_NAME,
        "mimeType": "application/vnd.google-apps.folder",
    }
    folder = svc.files().create(body=meta, fields="id").execute()
    fid = folder["id"]

    # 設定されていればオーナーのGmailアカウントに共有
    owner_email = os.environ.get("OWNER_EMAIL")
    if owner_email:
        svc.permissions().create(
            fileId=fid,
            body={"type": "user", "role": "writer", "emailAddress": owner_email},
            sendNotificationEmail=False,
        ).execute()

    _root_id_cache = fid
    return _root_id_cache


def _get_or_create_subfolder(slug: str) -> str:
    svc = _drive()
    parent = _root_id()
    safe = slug.replace("'", "\\'")
    q = (
        f"name='{safe}'"
        " and mimeType='application/vnd.google-apps.folder'"
        f" and '{parent}' in parents"
        " and trashed=false"
    )
    res = svc.files().list(q=q, fields="files(id)", pageSize=1).execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    folder = svc.files().create(
        body={"name": slug, "mimeType": "application/vnd.google-apps.folder", "parents": [parent]},
        fields="id",
    ).execute()
    return folder["id"]


def _list_files_in_folder(folder_id: str) -> list:
    svc = _drive()
    res = svc.files().list(
        q=f"'{folder_id}' in parents and trashed=false and mimeType contains 'image/'",
        fields="files(id, name, mimeType)",
        orderBy="createdTime",
        pageSize=200,
    ).execute()
    return res.get("files", [])


# ===== レシピ一覧 =====

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


# ===== API エンドポイント =====

@app.get("/api/recipes")
def list_recipes():
    svc = _drive()
    parent = _root_id()
    # ルート直下のサブフォルダとその写真数を一括取得
    q = f"mimeType='application/vnd.google-apps.folder' and '{parent}' in parents and trashed=false"
    res = svc.files().list(q=q, fields="files(id, name)", pageSize=200).execute()
    folder_map = {f["name"]: f["id"] for f in res.get("files", [])}

    counts: dict = {}
    for fname, fid in folder_map.items():
        r2 = svc.files().list(
            q=f"'{fid}' in parents and trashed=false and mimeType contains 'image/'",
            fields="files(id)",
            pageSize=1000,
        ).execute()
        counts[fname] = len(r2.get("files", []))

    return [
        {"name": r, "slug": _slug(r), "photo_count": counts.get(_slug(r), 0)}
        for r in _get_recipes()
    ]


@app.get("/api/photos/{slug}")
def list_photos(slug: str):
    folder_id = _get_or_create_subfolder(slug)
    files = _list_files_in_folder(folder_id)
    return [{"file_id": f["id"], "filename": f["name"]} for f in files]


@app.get("/api/photo/{file_id}")
def serve_photo(file_id: str):
    """Google Drive の画像をプロキシ配信（CORS回避）。"""
    svc = _drive()
    try:
        meta = svc.files().get(fileId=file_id, fields="mimeType").execute()
        content = svc.files().get_media(fileId=file_id).execute()
        return Response(content=content, media_type=meta.get("mimeType", "image/jpeg"))
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"ファイルが見つかりません: {e}")


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

    folder_id = _get_or_create_subfolder(slug)
    content = await file.read()

    from googleapiclient.http import MediaIoBaseUpload
    svc = _drive()
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=content_type)
    created = svc.files().create(
        body={"name": file.filename, "parents": [folder_id]},
        media_body=media,
        fields="id, name",
    ).execute()
    return {"file_id": created["id"], "filename": created["name"]}


@app.delete("/api/photos/{file_id}")
def delete_photo(file_id: str):
    svc = _drive()
    try:
        svc.files().delete(fileId=file_id).execute()
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"削除エラー: {e}")
    return {"result": "ok"}
