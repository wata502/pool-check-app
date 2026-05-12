"""
pool_writer.py
プール付属電気設備点検システム - PC側 Firebase → Excel 自動書き込みツール
Windows .exe として配布（PyInstaller でビルド）

既存の InspectionWriter と同じアーキテクチャで構築:
  - Firebase Realtime Database の /pool_inspections をリアルタイム監視（SSE）+ ポーリング
  - 新データ検知 → pool_mapping.json 参照 → win32com で Excel 書き込み
  - tkinter GUI（設定・ログ表示）
"""

import copy
import os
import sys
import json
import time
import math
import re
import logging
import logging.handlers
import threading
import queue
import traceback
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

try:
    import requests
    # ── HTTP 堅牢化のため Session + Retry を導入（urllib3 標準同梱）──
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
except ImportError:
    requests = None
    HTTPAdapter = None
    Retry = None

try:
    import sseclient
    SSE_AVAILABLE = True
except ImportError:
    sseclient = None
    SSE_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    import pywintypes
    # win32process は Excel ゾンビプロセスの PID 取得に使用（強制終了の最終手段）
    import win32process
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    win32process = None

# subprocess は taskkill /F による Excel プロセス強制終了に使用（_safe_com_cleanup 参照）
import subprocess

# ===== パス設定 =====
APP_NAME = "PoolWriter"
_APP_DIR = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
CONFIG_FILE = _APP_DIR / "pool_settings.json"
LOG_FILE    = _APP_DIR / "pool_writer.log"
MAPPING_FILE = _APP_DIR / "pool_mapping.json"

# ===== デフォルト設定 =====
DEFAULT_CONFIG = {
    "firebase_url": "",
    "firebase_api_key": "",
    "firebase_email": "",
    "firebase_password": "",
    "excel_folder": "",
    "auto_start": True,
    "version": "1.0.0",
}

# ===== ログ設定 =====
def setup_logging():
    file_fmt = logging.Formatter(fmt="%(asctime)s [%(levelname)s] %(message)s")
    console_fmt = logging.Formatter(fmt="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
    # RotatingFileHandler: 5MB x 3世代 (最大20MB) で自動ローテーション
    file_handler = logging.handlers.RotatingFileHandler(
        LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"
    )
    file_handler.setFormatter(file_fmt)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(console_fmt)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    for h in root.handlers[:]:
        root.removeHandler(h)
    root.addHandler(file_handler)
    root.addHandler(console_handler)

# ===== 設定の読み書き =====
def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                if k not in cfg:
                    cfg[k] = copy.deepcopy(v)
            return cfg
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

# ===== マッピング読み込み =====
def load_mapping():
    if MAPPING_FILE.exists():
        try:
            with open(MAPPING_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {k: v for k, v in data.items() if not k.startswith('_')}
        except Exception as e:
            logging.error(f"pool_mapping.json 読み込みエラー: {e}")
    return {}

# ===== Firebase 認証 =====
_fb_auth_lock     = threading.Lock()
_fb_auth_token    = None
_fb_auth_expiry   = 0.0
_fb_refresh_token = None
_fb_email         = ""
_fb_password      = ""

def _update_firebase_credentials(cfg):
    global _fb_email, _fb_password, _fb_auth_token, _fb_refresh_token, _fb_auth_expiry
    _fb_email    = cfg.get("firebase_email", "").strip()
    _fb_password = cfg.get("firebase_password", "").strip()
    with _fb_auth_lock:
        _fb_auth_token    = None
        _fb_refresh_token = None
        _fb_auth_expiry   = 0.0

def _get_firebase_token(api_key):
    global _fb_auth_token, _fb_auth_expiry, _fb_refresh_token
    if not api_key or not _fb_email or not _fb_password:
        return None
    with _fb_auth_lock:
        now = time.time()
        if _fb_auth_token and now < _fb_auth_expiry - 300:
            return _fb_auth_token
        if _fb_refresh_token:
            try:
                resp = requests.post(
                    f"https://securetoken.googleapis.com/v1/token?key={api_key}",
                    data=f"grant_type=refresh_token&refresh_token={_fb_refresh_token}",
                    headers={"Content-Type": "application/x-www-form-urlencoded"},
                    timeout=10)
                if resp.ok:
                    d = resp.json()
                    _fb_auth_token    = d.get("id_token")
                    _fb_refresh_token = d.get("refresh_token")
                    _fb_auth_expiry   = now + int(d.get("expires_in", 3600))
                    logging.info("[Auth] トークンをリフレッシュしました")
                    return _fb_auth_token
            except Exception as e:
                logging.warning(f"[Auth] リフレッシュ失敗: {e}")
        try:
            resp = requests.post(
                f"https://identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key={api_key}",
                json={"email": _fb_email, "password": _fb_password, "returnSecureToken": True},
                timeout=10)
            if resp.ok:
                d = resp.json()
                _fb_auth_token    = d.get("idToken")
                _fb_refresh_token = d.get("refreshToken")
                _fb_auth_expiry   = now + int(d.get("expiresIn", 3600))
                logging.info(f"[Auth] Firebase サインイン成功")
                return _fb_auth_token
            else:
                try:
                    fb_err = resp.json().get("error", {})
                    fb_msg = fb_err.get("message", f"HTTP {resp.status_code}")
                except Exception:
                    fb_msg = f"HTTP {resp.status_code}"
                logging.error(f"[Auth] サインイン失敗: {fb_msg}")
                raise RuntimeError(fb_msg)
        except RuntimeError:
            raise
        except Exception as e:
            logging.error(f"[Auth] サインインエラー: {e}")
            raise RuntimeError(str(e))

def _fb_auth_url(url, api_key):
    if not api_key:
        return url
    token = _get_firebase_token(api_key)
    if not token:
        return url
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}auth={token}"

# ===== HTTP 共通レイヤ（Firebase 通信の堅牢化） =====
# 目的:
#   - urllib3 Retry でステータス系（5xx/429）を自動リトライ
#   - SSL EOF / ConnectionError / ReadTimeout は別途、自前の指数バックオフで再試行
#     （SSLEOFError は urllib3 のリトライ対象に入りにくく、明示的なリトライが必要）
#   - 401/403 ではトークン強制再取得 → 1 回だけリトライ
#   - HTTP コネクションを Session で使い回し、TLS ハンドシェイク回数を削減（軽快動作）
def _build_http_session():
    if requests is None:
        return None
    s = requests.Session()
    if Retry is not None and HTTPAdapter is not None:
        retry = Retry(
            total=4, connect=4, read=4,
            status_forcelist=(429, 500, 502, 503, 504),
            backoff_factor=0.6,                 # 0.6,1.2,2.4,4.8 秒の指数バックオフ
            allowed_methods=frozenset(
                ["GET", "POST", "PATCH", "DELETE", "PUT"]),
            raise_on_status=False,
            respect_retry_after_header=True,
        )
        adapter = HTTPAdapter(
            max_retries=retry, pool_connections=8, pool_maxsize=16)
        s.mount("https://", adapter)
        s.mount("http://", adapter)
    return s

# プロセス全体で 1 つの Session を共有（短寿命/破棄時はスレッドセーフな再生成）
_HTTP_SESSION = _build_http_session()
_HTTP_SESSION_LOCK = threading.Lock()


def _http_request(method, url, *, api_key=None, json_body=None,
                  timeout=15, max_attempts=4, log_prefix="HTTP"):
    """
    Firebase 向けの堅牢な HTTP リクエスト。
      - SSL EOF / ConnectionError / ReadTimeout を最大 max_attempts 回まで指数バックオフ
      - 401/403 は token 強制再取得後に 1 度だけ追加リトライ
      - 戻り値: requests.Response（成功/最終応答）／全失敗時 None
    api_key を渡した場合、毎回 _fb_auth_url() で最新トークンを付与する
    （途中でトークンが更新されても古いトークンを焼き付けないため）。
    """
    if requests is None:
        return None
    global _HTTP_SESSION
    sess = _HTTP_SESSION
    if sess is None:
        with _HTTP_SESSION_LOCK:
            if _HTTP_SESSION is None:
                _HTTP_SESSION = _build_http_session()
            sess = _HTTP_SESSION

    last_err = None
    auth_retried = False
    for attempt in range(max_attempts):
        try:
            req_url = _fb_auth_url(url, api_key) if api_key else url
            resp = sess.request(
                method, req_url, json=json_body, timeout=timeout)
            # トークン期限切れ: 1 回だけ強制再認証して再試行
            if resp.status_code in (401, 403) and api_key and not auth_retried:
                auth_retried = True
                global _fb_auth_token, _fb_auth_expiry
                with _fb_auth_lock:
                    _fb_auth_token = None
                    _fb_auth_expiry = 0.0
                logging.warning(
                    f"[{log_prefix}] {resp.status_code} 認証失敗 — "
                    f"トークンを再取得して再試行")
                continue
            return resp
        except (requests.exceptions.SSLError,
                requests.exceptions.ConnectionError,
                requests.exceptions.ChunkedEncodingError,
                requests.exceptions.ReadTimeout) as e:
            # SSLEOFError や瞬断はここに来る — 指数バックオフで明示リトライ
            last_err = e
            if attempt >= max_attempts - 1:
                break
            wait = min(2 ** attempt, 20) + 0.25 * attempt
            logging.warning(
                f"[{log_prefix}] {type(e).__name__} 一時障害 "
                f"(試行{attempt+1}/{max_attempts}) — {wait:.1f}s 後リトライ")
            time.sleep(wait)
        except Exception as e:
            # その他は即時 break（呼び元で処理）
            last_err = e
            break
    if last_err is not None:
        logging.error(f"[{log_prefix}] 最終失敗 {method} {url.split('?')[0]}: {last_err}")
    return None


# ===== Firebase REST API =====
def firebase_get_pending(firebase_url, api_key=""):
    """pool_inspections から status=pending のレコードを全件取得してローカルフィルタリング"""
    base = firebase_url.rstrip("/") + "/pool_inspections.json"
    # _http_request 内で Retry / 401リトライ / 指数バックオフを実施するため、
    # 上位ではレスポンス検証のみに専念する（SSL EOF でクラッシュしないことを保証）
    resp = _http_request("GET", base, api_key=api_key, timeout=15,
                         log_prefix="FB-GET")
    if resp is None or not resp.ok:
        if resp is not None:
            logging.error(f"Firebase GET 失敗: HTTP {resp.status_code}")
        return {}
    try:
        data = resp.json()
    except Exception as e:
        logging.error(f"Firebase GET JSONパース失敗: {e}")
        return {}
    if not isinstance(data, dict):
        logging.info(
            f"firebase_get_pending: データなし (type={type(data).__name__})")
        return {}
    pending = {k: v for k, v in data.items()
               if isinstance(v, dict) and v.get("status") == "pending"}
    logging.info(
        f"firebase_get_pending: 取得総件数={len(data)}, pending件数={len(pending)}")
    return pending

def firebase_patch_status(firebase_url, record_id, patch_data, api_key=""):
    """レコードの status を更新（PATCH も _http_request 経由で堅牢化）"""
    url = firebase_url.rstrip("/") + f"/pool_inspections/{record_id}.json"
    resp = _http_request("PATCH", url, api_key=api_key, json_body=patch_data,
                         timeout=15, log_prefix="FB-PATCH")
    if resp is None:
        # 通信全断: ワーカー側で error カウントされる。冪等性を尊重し False を返却
        return False
    if not resp.ok:
        logging.error(
            f"Firebase PATCH 失敗 ({record_id}): HTTP {resp.status_code}")
        return False
    return True

# ===== Excel 書き込みユーティリティ =====
def col_to_num(col_str: str) -> int:
    result = 0
    for c in col_str.upper():
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result

def _set_cell(ws, row, col_letter, value):
    if value is None:
        safe_value = ""
    elif isinstance(value, bool):
        safe_value = value
    elif isinstance(value, (int, float)):
        safe_value = "" if (math.isnan(value) or math.isinf(value)) else value
    else:
        safe_value = str(value)
    ws.Cells(row, col_to_num(col_letter)).Value = safe_value

def _excel_pid(xl_app):
    """xl_app.Hwnd → GetWindowThreadProcessId で実 PID を取得。
    Quit() が COM 例外で失敗した場合のフォールバックに使用する。"""
    if xl_app is None or win32process is None:
        return None
    try:
        hwnd = xl_app.Hwnd
        if not hwnd:
            return None
        _tid, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid) if pid else None
    except Exception:
        return None


def _kill_pid_force(pid, timeout=3.0):
    """taskkill /F でプロセスを強制終了。CREATE_NO_WINDOW で黒窓抑止。"""
    if not pid:
        return
    try:
        subprocess.run(
            ["taskkill", "/F", "/PID", str(pid)],
            check=False, timeout=timeout,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=0x08000000,  # CREATE_NO_WINDOW
        )
    except Exception:
        pass


def _safe_com_cleanup(wb, xl_app):
    """COM オブジェクトの確実な解放。Quit() 失敗時は PID kill にフォールバック。
    ※ ゾンビ EXCEL.EXE が ~$ ロックファイルを握り続けて以後の書込が永続失敗する
       問題を防ぐため、解放経路を二段構えにしている。"""
    pid = _excel_pid(xl_app)  # Quit 前に PID を控えておく（Quit 後は Hwnd が無効になる）

    # ── ステップ 1: ワークブッククローズ（変更を破棄）──
    if wb is not None:
        try:
            wb.Saved = True   # 未保存ダイアログ抑止
        except Exception:
            pass
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass

    # ── ステップ 2: アプリ状態を正規モードに復帰してから Quit ──
    quit_ok = False
    if xl_app is not None:
        for prop, val in (("EnableEvents", True), ("DisplayAlerts", True),
                          ("ScreenUpdating", True)):
            try:
                setattr(xl_app, prop, val)
            except Exception:
                pass
        try:
            setattr(xl_app, "Calculation", -4105)  # xlCalculationAutomatic
        except Exception:
            pass
        try:
            xl_app.Quit()
            quit_ok = True
        except Exception:
            pass

    # ── ステップ 3: COM 参照解放（GC 待たず即座にプロセス解放を促す）──
    try:
        del wb
    except Exception:
        pass
    try:
        del xl_app
    except Exception:
        pass

    # ── ステップ 4: Quit が失敗 or 不完了の場合、PID を強制終了（最終手段）──
    if (not quit_ok) and pid:
        logging.warning(f"[COM] Excel.Quit 失敗 — PID {pid} を taskkill /F で強制終了")
        _kill_pid_force(pid)

# ===== 特記事項書込: 結合セル安全な Justify フォールバック =====
# 背景: テンプレートの A39:H49 は行ごとに A:H が結合されており、
#       Range("A:H").Justify() は MergeCells を含むと COM 例外
#       (-2146827284 "結合したセルには行えません") を投げる。
#       ここでは MergeCells を事前検出して Justify を回避し、
#       ColumnWidth 合算 × 全角/半角文字幅で手動ラップする。

def _is_range_merged(rng):
    """Range が結合セルを含むか。MergeCells が None（一部結合）の場合も True 扱い。"""
    try:
        m = rng.MergeCells
    except Exception:
        return False
    if m is None:
        return True
    return bool(m)


def _calc_row_capacity(ws, start_col, end_col, _row):
    """A:end_col の合計 ColumnWidth を「半角文字数の概算」として返す。
    Excel の ColumnWidth は標準フォントでの半角文字数に近い値。"""
    total = 0.0
    for c in range(col_to_num(start_col), col_to_num(end_col) + 1):
        try:
            total += float(ws.Columns(c).ColumnWidth or 0)
        except Exception:
            pass
    # 末尾余白マージン (5%) と最小値ガード（壊れたテンプレート対策）
    return max(int(total * 0.95), 12)


def _wrap_text_jp(text, max_units):
    """全角=2, 半角=1 で行に分割。空文字列は ['']。"""
    if not text:
        return [""]
    lines = []
    buf_chars = []
    units = 0
    for ch in text:
        # 0x7F 超えを全角扱い（簡易判定: 日本語運用では実用十分）
        w = 2 if ord(ch) > 0x7F else 1
        if units + w > max_units and buf_chars:
            lines.append("".join(buf_chars))
            buf_chars = [ch]
            units = w
        else:
            buf_chars.append(ch)
            units += w
    if buf_chars:
        lines.append("".join(buf_chars))
    return lines or [""]


def _write_paragraph_safely(ws, xl_app, start_col, end_col,
                             current_row, end_row, paragraph):
    """段落 1 つを current_row..end_row 範囲に書込み、消費行数を返す。
       戻り値: (consumed_rows, overflowed: bool)
       - 結合セル時は Justify を呼ばず手動ラップ
       - 非結合時は Justify を試行 → COM例外時は手動ラップへフォールバック
       - 範囲超過時は最終セルに残りを連結（情報欠落防止）"""
    first_cell = ws.Range(f"{start_col}{current_row}")

    # ── 結合セル判定: 結合なら最初から手動ラップ（Justify を呼ばない）──
    if _is_range_merged(first_cell):
        capacity = _calc_row_capacity(ws, start_col, end_col, current_row)
        chunks = _wrap_text_jp(paragraph, capacity)
        consumed = 0
        for i, chunk in enumerate(chunks):
            r = current_row + i
            if r > end_row:
                # 範囲オーバー: 最終セルに残りを連結（情報欠落の防止）
                try:
                    last = ws.Range(f"{start_col}{end_row}")
                    existing = str(last.Value or "")
                    rest = "".join(chunks[i:])
                    last.Value = (existing + " " + rest).strip()
                except Exception:
                    pass
                return (max(consumed, 1), True)
            try:
                ws.Range(f"{start_col}{r}").Value = chunk
            except Exception as e:
                logging.warning(f"  特記事項 結合セル書込失敗 (行{r}): {e}")
            consumed = i + 1
        return (max(consumed, 1), False)

    # ── 非結合: 一旦書込 → Justify を試行 ──
    try:
        first_cell.Value = paragraph
    except Exception as e:
        logging.warning(f"  特記事項 値書込失敗 (行{current_row}): {e}")
        return (1, False)

    prev_alerts = xl_app.DisplayAlerts
    used_justify = False
    try:
        xl_app.DisplayAlerts = True
        ws.Range(
            f"{start_col}{current_row}:{end_col}{end_row}"
        ).Justify()
        used_justify = True
    except Exception as e:
        # COM 例外（結合セル混在含む）時は値を一旦消して手動ラップへ
        logging.info(f"  Justify 失敗 → 手動ラップへフォールバック (行{current_row}): {e}")
        try:
            first_cell.Value = ""
        except Exception:
            pass
    finally:
        xl_app.DisplayAlerts = prev_alerts

    if not used_justify:
        # 手動ラップ実行
        capacity = _calc_row_capacity(ws, start_col, end_col, current_row)
        chunks = _wrap_text_jp(paragraph, capacity)
        consumed = 0
        for i, chunk in enumerate(chunks):
            r = current_row + i
            if r > end_row:
                try:
                    last = ws.Range(f"{start_col}{end_row}")
                    existing = str(last.Value or "")
                    rest = "".join(chunks[i:])
                    last.Value = (existing + " " + rest).strip()
                except Exception:
                    pass
                return (max(consumed, 1), True)
            try:
                ws.Range(f"{start_col}{r}").Value = chunk
            except Exception:
                pass
            consumed = i + 1
        return (max(consumed, 1), False)

    # Justify 成功時: 消費行数を A 列の値で再検出
    used = 1
    for r in range(current_row + 1, end_row + 1):
        try:
            v = ws.Range(f"{start_col}{r}").Value
        except Exception:
            v = None
        if v is not None and str(v).strip():
            used = r - current_row + 1
        else:
            break
    return (used, False)


# ===== プール点検 Excel 書き込み =====
def write_pool_record(excel_path: str, record: dict, school_info: dict) -> tuple[bool, str]:
    """
    pool_mapping.json の情報に従い、プール点検データを Excel に書き込む。
    戻り値: (success, message)
    """
    if not WIN32_AVAILABLE:
        return False, "pywin32 がインストールされていません"

    abs_path = os.path.normpath(str(Path(excel_path).resolve()))
    if not os.path.exists(abs_path):
        return False, f"ファイルが見つかりません: {abs_path}"

    lock_file = Path(abs_path).parent / ("~$" + Path(abs_path).name)
    if lock_file.exists():
        return False, f"LOCKED:Excelが開かれています: {Path(abs_path).name}"

    pythoncom.CoInitialize()
    xl_app = None
    wb = None

    try:
        xl_app = win32com.client.DispatchEx("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        xl_app.ScreenUpdating = False
        xl_app.EnableEvents = False

        wb = xl_app.Workbooks.Open(abs_path)
        xl_app.Calculation = -4135  # xlCalculationManual

        # シート特定
        sheet_name = school_info.get("sheet_name", "")
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if not ws and (
                wb.Sheets(i).Name == sheet_name or
                (not sheet_name and i == 1)
            ):
                ws = wb.Sheets(i)

        if ws is None:
            ws = wb.Sheets(1)

        logging.info(f"  シート: {ws.Name}")

        header_map = school_info.get("header", {}) or {}
        header_data = record.get("header") or {}  # Firebase が null を返した場合も空dictとして扱う

        # ===== 基本情報 =====
        def write_header(cell_key, value):
            cell = header_map.get(cell_key)
            if cell and value is not None and value != "":
                col = re.sub(r'\d', '', cell).upper()
                row = int(re.sub(r'[A-Za-z]', '', cell))
                _set_cell(ws, row, col, str(value))

        write_header("date_cell",        header_data.get("date", ""))
        write_header("time_cell",        header_data.get("time", ""))
        write_header("weather_cell",     header_data.get("weather", ""))
        write_header("temperature_cell", header_data.get("temperature"))
        write_header("note_cell",        header_data.get("note", ""))

        # 湿度: Excelのセルはパーセント書式のため、数値を÷100して書き込む
        # 例: 73 → 0.73 → セルに "73%" と表示される（str(73)で書くと "7300%" になるバグを防ぐ）
        humidity_raw = header_data.get("humidity")
        if humidity_raw not in (None, ""):
            humidity_cell = header_map.get("humidity_cell")
            if humidity_cell:
                hum_col = re.sub(r'\d', '', humidity_cell).upper()
                hum_row = int(re.sub(r'[A-Za-z]', '', humidity_cell))
                try:
                    hum_float = float(str(humidity_raw))   # ① まず float 変換（失敗時は except へ）
                    _set_cell(ws, hum_row, hum_col, hum_float / 100.0)  # ② ÷100 してパーセントセルへ
                except (ValueError, TypeError):
                    # None・空文字以外の非数値（"不明" 等）が来た場合: クラッシュさせずスキップ
                    logging.warning(f"  湿度の数値変換に失敗しました（値: {humidity_raw!r}）— 書き込みをスキップします")

        logging.info("  基本情報書き込み完了")

        # ===== 絶縁抵抗測定 =====
        insulation_list = record.get("insulation") or []  # None/null → 空リスト
        insulation_cols = school_info.get("insulation_cols",
                          {"value": "D", "judgment": "F", "remarks": "G"})
        ins_by_row = {}
        if isinstance(insulation_list, list):
            for item in insulation_list:
                if isinstance(item, dict) and "row" in item:
                    ins_by_row[int(item["row"])] = item
        elif isinstance(insulation_list, dict):
            # レガシー dict フォーマット対応: {row: {...}, ...}
            for row_key, item in insulation_list.items():
                if isinstance(item, dict):
                    ins_by_row[int(row_key)] = item
        for row_info in school_info.get("insulation", []):
            row = row_info["row"]
            entry = ins_by_row.get(row)
            if not entry:
                continue
            if entry.get("value") not in (None, ""):
                try:
                    _set_cell(ws, row, insulation_cols["value"], float(entry["value"]))
                except (ValueError, TypeError):
                    # "-"（測定不可）等の非数値文字列はそのまま書き込み
                    _set_cell(ws, row, insulation_cols["value"], entry["value"])
            # 判定: "-"（測定不可）も明示値として書き込む（除外しない）
            if entry.get("judgment") not in (None, ""):
                _set_cell(ws, row, insulation_cols["judgment"], entry["judgment"])
            # 備考: 3状態分岐
            #   ""           → 上書きしない（前回値温存）
            #   空白文字のみ → 明示クリア（Excelセルを空に）
            #   通常文字     → 上書き
            rem = entry.get("remarks", "") or ""
            if rem != "":
                if rem.strip() == "":
                    _set_cell(ws, row, insulation_cols["remarks"], "")
                else:
                    _set_cell(ws, row, insulation_cols["remarks"], rem)

        logging.info("  絶縁抵抗測定書き込み完了")

        # ===== 接地抵抗測定 =====
        grounding_list = record.get("grounding") or []  # None/null → 空リスト
        grounding_cols = school_info.get("grounding_cols",
                         {"value": "D", "judgment": "F", "remarks": "G"})
        method_cell = school_info.get("grounding_method_cell", "H23")
        grounding_method = record.get("grounding_method", "")
        if grounding_method:
            mc = re.sub(r'\d', '', method_cell).upper()
            mr = int(re.sub(r'[A-Za-z]', '', method_cell))
            _set_cell(ws, mr, mc, grounding_method)

        gnd_by_row = {}
        if isinstance(grounding_list, list):
            for item in grounding_list:
                if isinstance(item, dict) and "row" in item:
                    gnd_by_row[int(item["row"])] = item
        elif isinstance(grounding_list, dict):
            # レガシー dict フォーマット対応
            for row_key, item in grounding_list.items():
                if isinstance(item, dict):
                    gnd_by_row[int(row_key)] = item
        for row_info in school_info.get("grounding", []):
            row = row_info["row"]
            entry = gnd_by_row.get(row)
            if not entry:
                continue
            if entry.get("value") not in (None, ""):
                try:
                    _set_cell(ws, row, grounding_cols["value"], float(entry["value"]))
                except (ValueError, TypeError):
                    # "-"（測定不可）等の非数値文字列はそのまま書き込み
                    _set_cell(ws, row, grounding_cols["value"], entry["value"])
            # 判定: "-"（測定不可）も明示値として書き込む（除外しない）
            if entry.get("judgment") not in (None, ""):
                _set_cell(ws, row, grounding_cols["judgment"], entry["judgment"])
            # 備考: 3状態分岐（空=温存／空白のみ=クリア／通常=上書き）
            rem = entry.get("remarks", "") or ""
            if rem != "":
                if rem.strip() == "":
                    _set_cell(ws, row, grounding_cols["remarks"], "")
                else:
                    _set_cell(ws, row, grounding_cols["remarks"], rem)

        logging.info("  接地抵抗測定書き込み完了")

        # ===== 漏電遮断器テスト =====
        breaker_list = record.get("breaker_test") or []  # None/null → 空リスト
        breaker_cols = school_info.get("breaker_test_cols",
                       {"judgment": "F", "remarks": "G"})
        brk_by_row = {}
        if isinstance(breaker_list, list):
            for item in breaker_list:
                if isinstance(item, dict) and "row" in item:
                    brk_by_row[int(item["row"])] = item
        elif isinstance(breaker_list, dict):
            # レガシー dict フォーマット対応
            for row_key, item in breaker_list.items():
                if isinstance(item, dict):
                    brk_by_row[int(row_key)] = item
        for row_info in school_info.get("breaker_test", []):
            row = row_info["row"]
            entry = brk_by_row.get(row)
            if not entry:
                continue
            # 判定: "-"（測定不可）も明示値として書き込む（除外しない）
            if entry.get("judgment") not in (None, ""):
                _set_cell(ws, row, breaker_cols["judgment"], entry["judgment"])
            # 備考: 3状態分岐（空=温存／空白のみ=クリア／通常=上書き）
            rem = entry.get("remarks", "") or ""
            if rem != "":
                if rem.strip() == "":
                    _set_cell(ws, row, breaker_cols["remarks"], "")
                else:
                    _set_cell(ws, row, breaker_cols["remarks"], rem)

        logging.info("  漏電遮断器テスト書き込み完了")

        # ===== 特記事項 =====
        # 仕様:
        #  - ユーザーの入力改行（\n）を段落区切りとして尊重する
        #  - 各段落を順に書き込み、その段落について Justify() を適用して
        #    H列幅で機械的に折返す（I列以降にはみ出さない）
        #  - 次段落は、前段落が Justify で消費した最終行の次から書く
        #  - 範囲(end_row)を超える段落は、最終セルに連結してオーバーフロー回避
        # 特記事項: 3状態分岐
        #   ""           → 上書きしない（前回値温存）
        #   空白文字のみ → 明示クリア（範囲を空にしてリターン）
        #   通常文字     → 既存セルをクリアしたうえで段落を書き込み
        special_notes = record.get("special_notes") or ""  # None/null → 空文字列（None.strip()クラッシュを防ぐ）
        if special_notes != "":
            notes_range = school_info.get("special_notes_range",
                          {"start": "A39", "end": "H49"})
            start_cell = notes_range.get("start", "A39")
            end_cell   = notes_range.get("end",   "H49")
            start_col  = re.sub(r'\d', '', start_cell).upper()
            start_row  = int(re.sub(r'[A-Za-z]', '', start_cell))
            end_col    = re.sub(r'\d', '', end_cell).upper()
            end_row    = int(re.sub(r'[A-Za-z]', '', end_cell))

            # 仕様（重要）:
            #  - ユーザーは手動運用で「文字の割付」(= Range.Justify メソッド) を
            #    A:H 範囲に対して使用している（I 列まで広げると印刷時にはみ出す）。
            #  - したがって書込み側もテンプレートのセル書式・結合状態・行高・
            #    WrapText 設定を一切変更してはならない。値の書込みと Justify 呼出
            #    のみを行う。
            #  - 列範囲はマッピング側の end_col に従う（特記事項は通常 H49）。
            justify_end_col = end_col  # マッピング設定値をそのまま使用（=H）

            # 各行の値だけクリア（書式・結合・WrapText・行高は触らない）
            # 注: 空白文字のみの入力＝明示クリアの場合、ここで終了する
            for r in range(start_row, end_row + 1):
                try:
                    ws.Range(f"{start_col}{r}").Value = ""
                except Exception:
                    pass

            # 空白文字のみ＝明示クリア。段落書き込みはスキップ
            if not special_notes.strip():
                logging.info("  特記事項クリア完了（明示クリア）")
            else:
                # 改行コードを統一して段落分割（CR/LF, CR, LF いずれにも対応）
                paragraphs = (special_notes
                              .replace("\r\n", "\n")
                              .replace("\r", "\n")
                              .split("\n"))

                current_row = start_row
                overflow_logged = False
                for para in paragraphs:
                    # 空段落はそのまま空行として1行送る（連続改行も尊重）
                    if not para.strip():
                        if current_row <= end_row:
                            current_row += 1
                        continue

                    # 範囲オーバーフロー: 最終セルに残りを連結（情報欠落の防止）
                    if current_row > end_row:
                        try:
                            last_cell = ws.Range(f"{start_col}{end_row}")
                            existing = str(last_cell.Value or "")
                            last_cell.Value = (existing + " " + para).strip()
                        except Exception:
                            pass
                        if not overflow_logged:
                            logging.warning("  特記事項範囲超過 — 最終セルに連結")
                            overflow_logged = True
                        continue

                    # 結合セル安全な書込ヘルパーへ委譲
                    #   - 結合セルなら手動ラップ（Justify を呼ばない＝COM 例外を完全回避）
                    #   - 非結合なら Justify を試行 → 失敗時は手動ラップへフォールバック
                    used_rows, overflowed = _write_paragraph_safely(
                        ws, xl_app,
                        start_col, justify_end_col,
                        current_row, end_row, para)
                    current_row += used_rows
                    if overflowed and not overflow_logged:
                        logging.warning("  特記事項範囲超過 — 最終セルに連結")
                        overflow_logged = True

                # フォント・行高・WrapText・結合状態などのセル書式は
                # テンプレートの設定を尊重し、書込み側からは変更しない。
                logging.info("  特記事項書き込み完了")

        # ===== 保存 =====
        xl_app.Calculation = -4105  # xlCalculationAutomatic
        wb.Save()
        logging.info("  保存完了")
        # 旧コード: time.sleep(1.0) — 確実な保存待ちのつもりだが、Save() は同期呼出のため不要。
        # 削除して 1 件あたり 1 秒のレイテンシを削減（軽快動作）。

        wb.Saved = True
        wb.Close()
        xl_app.Quit()
        wb = None
        xl_app = None

        return True, f"書き込み完了 → {os.path.basename(abs_path)}"

    except Exception as e:
        tb = traceback.format_exc()
        logging.error(f"Excel書き込み例外:\n{tb}")
        tb_lines = [l for l in tb.strip().splitlines() if l.strip()]
        loc = tb_lines[-2] if len(tb_lines) >= 2 else ""
        _safe_com_cleanup(wb, xl_app)
        wb = None
        xl_app = None
        return False, f"Excel操作失敗: {e}  (位置: {loc.strip()})"
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def find_excel_file(base_folder: str, file_name: str) -> str | None:
    """フォルダ内を再帰検索してファイルを探す"""
    base = Path(base_folder)
    if not base.is_dir():
        return None
    for p in base.rglob(file_name):
        if p.is_file():
            return str(p.resolve())
    return None


# ===== キュー・ワーカー =====
_task_queue      = queue.Queue()
_queued_ids      = set()
_queued_ids_lock = threading.Lock()
_recently_written = {}       # {record_id: {"updatedAt": str, "written_at": float}}
_IDEMPOTENCY_SEC = 300
_sse_connected   = threading.Event()   # SSE接続状態フラグ

def enqueue_record(record_id):
    with _queued_ids_lock:
        if record_id in _queued_ids:
            return False
        _queued_ids.add(record_id)
    _task_queue.put(record_id)
    return True


def _is_newer_than_cached(record_id, record_data):
    """
    PoolPoller（CMD 受信時／定期ポーリング時）の冪等性判定ヘルパー。
    ワーカー側 _process_one の「updatedAt ベース冪等性」と整合させる。

    True（=エンキュー対象） を返す条件:
        1. キャッシュ無し（初回）
        2. キャッシュから _IDEMPOTENCY_SEC 経過（時間経過で再試行）
        3. 冪等期間内でも、取得したレコードの updatedAt が
           キャッシュ済み updatedAt より新しい（＝ユーザーによる正当な更新）

    False を返す条件:
        - 冪等期間内で、かつ updatedAt がキャッシュと同じ／古い
          （多重キューイング・同一データの再投入を防止）

    引数:
        record_id   : Firebase 上のレコード ID
        record_data : firebase_get_pending() が返す辞書値（dict）。
                      None や updatedAt キー欠落も許容する。
    """
    cached = _recently_written.get(record_id)
    if not cached:
        return True
    # 冪等期間を過ぎていれば無条件で投入可
    if (time.time() - cached.get("written_at", 0)) >= _IDEMPOTENCY_SEC:
        return True
    # 冪等期間内でも updatedAt が進んでいれば「正当な再送信」として投入する
    live_updated_at = ""
    if isinstance(record_data, dict):
        live_updated_at = record_data.get("updatedAt", "") or ""
    cached_updated_at = cached.get("updatedAt", "") or ""
    if live_updated_at and live_updated_at > cached_updated_at:
        return True
    return False


class PoolWriteWorker:
    def __init__(self, cfg, mapping_data, log_callback=None, error_callback=None):
        self.cfg = cfg
        self.mapping_data = mapping_data
        self.log_callback = log_callback
        self.error_callback = error_callback
        self._thread = None
        self._stop_event = threading.Event()
        self._write_count = 0
        self._error_count = 0
        self._skip_count  = 0

    @property
    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.is_running:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._worker_loop, daemon=True, name="Pool-Excel-Worker")
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        _task_queue.put(None)

    def reload_mapping(self, mapping_data):
        self.mapping_data = mapping_data

    def get_stats(self):
        return {"write": self._write_count, "error": self._error_count,
                "skip": self._skip_count, "queue": _task_queue.qsize()}

    def _log(self, msg):
        logging.info(msg)
        if self.log_callback:
            self.log_callback(msg)

    def _worker_loop(self):
        try:
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            self._log("[Worker] プール点検 Excel書き込みワーカー開始")
            ses_w = ses_e = ses_s = 0

            while not self._stop_event.is_set():
                try:
                    record_id = _task_queue.get(timeout=2)
                except queue.Empty:
                    if ses_w + ses_e + ses_s > 0:
                        self._log(f"=== 一括完了: 成功={ses_w} 失敗={ses_e} スキップ={ses_s} ===")
                        ses_w = ses_e = ses_s = 0
                    # 古い冪等性キャッシュを定期クリーンアップ（PW-3）
                    now = time.time()
                    stale = [k for k, v in _recently_written.items()
                             if now - v["written_at"] >= _IDEMPOTENCY_SEC]
                    for k in stale:
                        del _recently_written[k]
                    continue

                if record_id is None:
                    _task_queue.task_done()
                    break

                pw, pe, ps = self._write_count, self._error_count, self._skip_count
                try:
                    self._process_one(record_id)
                except Exception as e:
                    self._log(f"[Worker] 予期しないエラー ({record_id}): {e}")
                finally:
                    _task_queue.task_done()
                    # ③ SSE競合防止: 処理完了後に _queued_ids から除去
                    #    （処理中は _queued_ids に残るため、ポーリングによる重複投入を防ぐ）
                    with _queued_ids_lock:
                        _queued_ids.discard(record_id)

                ses_w += self._write_count - pw
                ses_e += self._error_count - pe
                ses_s += self._skip_count  - ps

            self._log("[Worker] Excel書き込みワーカー終了")
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _process_one(self, record_id):
        firebase_url = self.cfg.get("firebase_url", "").strip()
        api_key      = self.cfg.get("firebase_api_key", "").strip()
        excel_folder = self.cfg.get("excel_folder", "").strip()

        # Firebase からレコード取得
        record = self._fetch_record(record_id)
        if not record:
            self._log(f"  [Skip] レコードなし: {record_id}")
            return

        if record.get("status") != "pending":
            self._skip_count += 1
            return

        # 冪等性チェック（updatedAt ベース）
        live_updated_at = record.get("updatedAt", "")
        cached = _recently_written.get(record_id)
        if cached and (time.time() - cached["written_at"]) < _IDEMPOTENCY_SEC:
            cached_updated_at = cached["updatedAt"]
            if live_updated_at and live_updated_at <= cached_updated_at:
                self._log(f"  👻 重複スキップ: {record_id}")
                self._skip_count += 1
                return

        school_id = record.get("school_id", "")
        self._log(f"  📝 処理中: {record_id} / 学校ID={school_id}")

        # マッピング検索
        school_info = self.mapping_data.get(school_id)
        if not school_info:
            err = f"school_id '{school_id}' のマッピングが見つかりません"
            self._log(f"  ✗ {err}")
            if firebase_url:
                firebase_patch_status(firebase_url, record_id,
                    {"status": "error", "error_msg": err}, api_key=api_key)
            self._error_count += 1
            return

        # ※「プール運用なし」でも日付等の基本情報だけを書き込む運用のため、has_data チェックは行わない
        #   Noneガード処理により数値データが空でもクラッシュせず、基本情報のみ書き込んで「成功」になる

        # Excel ファイル探索
        file_name = school_info.get("file_name", "")
        if not file_name:
            err = f"ファイル名が未設定: {school_id}"
            self._log(f"  ✗ {err}")
            if firebase_url:
                firebase_patch_status(firebase_url, record_id,
                    {"status": "skipped", "skip_reason": err}, api_key=api_key)
            self._skip_count += 1
            return

        if not excel_folder or not os.path.isdir(excel_folder):
            err = "Excelフォルダが設定されていないか存在しません"
            self._log(f"  ✗ {err}")
            if firebase_url:
                firebase_patch_status(firebase_url, record_id,
                    {"status": "error", "error_msg": err}, api_key=api_key)
            self._error_count += 1
            return

        excel_path = find_excel_file(excel_folder, file_name)
        if not excel_path:
            for ext in (".xlsm", ".xlsx"):
                alt = file_name.rsplit(".", 1)[0] + ext
                excel_path = find_excel_file(excel_folder, alt)
                if excel_path:
                    break

        if not excel_path:
            err = f"Excelファイルが見つかりません: {file_name}"
            self._log(f"  ✗ {err}")
            if firebase_url:
                firebase_patch_status(firebase_url, record_id,
                    {"status": "error", "error_msg": err}, api_key=api_key)
            if self.error_callback:
                self.error_callback(school_id, err)
            self._error_count += 1
            return

        self._log(f"  → {excel_path}")

        # Excel 書き込み（リトライ付き）
        ok, msg = False, ""
        for attempt in range(3):
            ok, msg = write_pool_record(excel_path, record, school_info)
            if ok:
                break
            if "LOCKED:" in msg and attempt < 2:
                self._log(f"  ⚠ ロック検出 (試行{attempt+1}/3) — 3秒待機...")
                time.sleep(3)
            else:
                break

        if ok:
            self._log(f"  ✓ {msg}")
            # 冪等性キャッシュに updatedAt 付きで記録
            _recently_written[record_id] = {
                "updatedAt": live_updated_at,
                "written_at": time.time(),
            }
            firebase_patch_status(firebase_url, record_id,
                {"status": "success", "syncAt": datetime.now().isoformat()},
                api_key=api_key)
            self._log(f"  📡 status=success を送信")
            self._write_count += 1
        else:
            self._log(f"  ✗ {msg}")
            firebase_patch_status(firebase_url, record_id,
                {"status": "error", "error_msg": msg}, api_key=api_key)
            if self.error_callback:
                self.error_callback(school_id, msg)
            self._error_count += 1

    def _fetch_record(self, record_id):
        # _http_request 経由で SSL EOF / 401 をリトライ。クラッシュさせない。
        firebase_url = self.cfg.get("firebase_url", "").strip()
        api_key = self.cfg.get("firebase_api_key", "").strip()
        if not firebase_url:
            return None
        url = firebase_url.rstrip("/") + f"/pool_inspections/{record_id}.json"
        resp = _http_request("GET", url, api_key=api_key, timeout=10,
                             log_prefix="FB-FETCH")
        if resp is None or resp.status_code != 200:
            if resp is not None:
                self._log(f"  ⚠ レコード取得失敗 ({record_id}): HTTP {resp.status_code}")
            return None
        try:
            data = resp.json()
            if isinstance(data, dict):
                return data
        except Exception as e:
            self._log(f"  ⚠ レコード JSON パース失敗 ({record_id}): {e}")
        return None


class PoolPoller:
    """5秒間隔で Firebase をポーリングして pending レコード + pool_commands を検知する"""
    POLL_INTERVAL = 5

    def __init__(self, cfg, log_callback=None):
        self.cfg = cfg
        self.log_callback = log_callback
        self._thread = None
        self._stop_event = threading.Event()
        self._poll_count = 0
        self._last_cmd_ts = None

    @property
    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.is_running:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._poll_loop, daemon=True, name="Pool-Poller")
        self._thread.start()

    def stop(self):
        self._stop_event.set()

    def _log(self, msg):
        logging.info(msg)
        if self.log_callback:
            self.log_callback(msg)

    def _poll_loop(self):
        while not self._stop_event.is_set():
            try:
                firebase_url = self.cfg.get("firebase_url", "").strip()
                if firebase_url:
                    api_key = self.cfg.get("firebase_api_key", "").strip()

                    # ── 1. pool_commands/run_now を確認（スマホの「PC書込」ボタン用）──
                    # _http_request 経由で SSL EOF / 瞬断は内部リトライ済み。
                    # それでも全失敗した場合のみここに到達する（loop 全体の継続性保護）。
                    try:
                        cmd_url = firebase_url.rstrip("/") + "/pool_commands/run_now.json"
                        # CMD は ~ 5 秒監視が要件のため short timeout は維持しつつ、
                        # _http_request 内のバックオフで瞬断を吸収する
                        resp = _http_request("GET", cmd_url, api_key=api_key,
                                             timeout=8, max_attempts=2,
                                             log_prefix="CMD-GET")
                        if resp is not None and resp.status_code == 200:
                            try:
                                data = resp.json()
                            except Exception:
                                data = None
                            if isinstance(data, dict) and data.get("requestedAt"):
                                ts = data["requestedAt"]
                                if ts != self._last_cmd_ts:
                                    self._last_cmd_ts = ts
                                    self._log("[CMD] 📱 スマホからPC書込指示を受信")
                                    # コマンド削除も _http_request 経由（失敗無視）
                                    _http_request("DELETE", cmd_url, api_key=api_key,
                                                  timeout=8, max_attempts=2,
                                                  log_prefix="CMD-DEL")
                                    # pending レコードをキューに投入
                                    records = firebase_get_pending(firebase_url, api_key)
                                    n = 0
                                    for rid, rec_data in records.items():
                                        if _is_newer_than_cached(rid, rec_data):
                                            if enqueue_record(rid):
                                                n += 1
                                    if n > 0:
                                        self._log(f"[CMD] {n} 件をキューに投入")
                                    else:
                                        self._log("[CMD] 未処理レコード: なし")
                    except Exception as e:
                        # 何らかの想定外例外でもポーリング自体は継続させる
                        logging.warning(f"[CMD] コマンドチェック想定外例外: {e}")

                    # ── 2. SSE 未接続時のポーリング（pending レコードの自動検出）──
                    #    冪等性は CMD 側と同じく updatedAt ベースで判定する
                    if not _sse_connected.is_set():
                        records = firebase_get_pending(firebase_url, api_key)
                        if records:
                            self._log(f"[Poll] 未処理 {len(records)} 件を検知")
                            for record_id, rec_data in records.items():
                                if _is_newer_than_cached(record_id, rec_data):
                                    enqueue_record(record_id)
            except Exception as e:
                logging.warning(f"[Poll] ポーリングエラー: {e}")
            self._stop_event.wait(self.POLL_INTERVAL)


class PoolSSEListener:
    """Firebase Realtime Database の /pool_inspections をSSEで監視"""
    READ_TIMEOUT = 120

    def __init__(self, cfg, log_callback=None, status_callback=None):
        self.cfg = cfg
        self.log_callback = log_callback
        self.status_callback = status_callback
        self._thread = None
        self._stop_event = threading.Event()
        self._response = None

    @property
    def is_running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.is_running:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(
            target=self._run_loop, daemon=True, name="Pool-SSE-Listener")
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        _sse_connected.clear()
        self._force_close_response()
        self._set_status("stopped", "SSE 停止")

    def _force_close_response(self):
        resp = self._response
        if not resp:
            return
        try:
            resp.raw.close()
        except Exception:
            pass
        try:
            resp.close()
        except Exception:
            pass

    def _cleanup_response(self):
        if self._response:
            try:
                self._response.close()
            except Exception:
                pass
            self._response = None

    def _log(self, msg):
        logging.info(msg)
        if self.log_callback:
            self.log_callback(msg)

    def _set_status(self, state, msg):
        if self.status_callback:
            self.status_callback(state, msg)

    def _run_loop(self):
        RECONNECT_BASE = 3
        RECONNECT_MAX  = 60
        wait = RECONNECT_BASE

        while not self._stop_event.is_set():
            firebase_url = self.cfg.get("firebase_url", "").strip()
            if not firebase_url:
                self._set_status("stopped", "Firebase URL 未設定")
                self._stop_event.wait(10)
                continue

            api_key = self.cfg.get("firebase_api_key", "").strip()
            sse_url = _fb_auth_url(firebase_url.rstrip("/") + "/pool_inspections.json", api_key)
            self._log(f"[SSE] 接続中: {firebase_url.rstrip('/')}/pool_inspections.json")
            self._set_status("reconnecting", "Firebase に接続中...")

            try:
                headers = {"Accept": "text/event-stream", "Cache-Control": "no-cache"}
                self._response = requests.get(
                    sse_url, stream=True, headers=headers,
                    timeout=(10, self.READ_TIMEOUT))
                self._response.raise_for_status()

                client = sseclient.SSEClient(self._response)
                self._set_status("connected", "🟢 リアルタイム監視中")
                self._log("[SSE] ✓ 接続成功 — リアルタイム監視を開始しました")
                wait = RECONNECT_BASE
                _sse_connected.set()

                # 接続復帰時: 未処理レコードをキューに投入
                self._log(f"[SSE] 接続復帰 — 未処理レコードを確認中...")
                try:
                    records = firebase_get_pending(firebase_url, api_key)
                    enqueued = 0
                    for rid in records:
                        if enqueue_record(rid):
                            enqueued += 1
                    if enqueued == 0:
                        self._log("  未処理レコード: なし")
                except Exception as _pe:
                    self._log(f"[SSE] 接続復帰エンキューエラー: {_pe}")

                for event in client.events():
                    if self._stop_event.is_set():
                        break

                    if event.event not in ("put", "patch"):
                        continue

                    try:
                        payload = json.loads(event.data)
                    except (json.JSONDecodeError, TypeError):
                        continue

                    path = payload.get("path", "")
                    data = payload.get("data")

                    if data is None:
                        continue

                    # ── path="/" → フルデータ（初回接続）──
                    if path == "/":
                        if isinstance(data, dict):
                            pending_ids = [k for k, v in data.items()
                                           if isinstance(v, dict)
                                           and v.get("status") == "pending"]
                            if pending_ids:
                                enqueued = 0
                                for rid in pending_ids:
                                    if enqueue_record(rid):
                                        enqueued += 1
                                if enqueued > 0:
                                    self._log(f"[SSE] 初回データ: {enqueued} 件をキューに投入")
                            else:
                                self._log("[SSE] 初回データ受信: 未処理レコードなし")
                        continue

                    # ── path="/{record_id}" → 個別レコード──
                    record_id = path.lstrip("/")
                    if "/" in record_id:
                        continue

                    # status が pending 以外は無視
                    if not isinstance(data, dict) or data.get("status") != "pending":
                        continue

                    if enqueue_record(record_id):
                        self._log(f"[SSE] 新規書込要求を受信: {record_id}")

            except Exception as e:
                _sse_connected.clear()
                self._cleanup_response()

                if self._stop_event.is_set():
                    break

                self._log(f"[SSE] 接続エラー: {type(e).__name__} — {wait}秒後に再接続します")

            self._set_status("reconnecting", f"再接続待ち ({wait}秒)...")
            self._stop_event.wait(wait)
            wait = min(wait * 2, RECONNECT_MAX)

        self._cleanup_response()
        self._log("[SSE] リスナースレッドを終了しました")


# ===== tkinter GUI =====
class PoolWriterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PoolWriter - プール点検 Excel書き込み")
        self.root.geometry("680x560")
        self.root.minsize(560, 400)

        self.cfg     = load_config()
        self.mapping = load_mapping()

        self.worker = None
        self.poller = None
        self.sse_listener = None
        self._tray_icon = None

        setup_logging()
        _update_firebase_credentials(self.cfg)

        self._build_ui()
        self.log(f"PoolWriter 起動 (マッピング: {len(self.mapping)} 校)")

        if not self.mapping:
            self.log("⚠ pool_mapping.json が見つかりません。"
                     "setup_mapping.py を実行してください。")

        if self.cfg.get("auto_start") and self.cfg.get("firebase_url"):
            self.root.after(800, self.start_monitoring)

        if TRAY_AVAILABLE:
            self.root.protocol("WM_DELETE_WINDOW", self._on_close_to_tray)
        else:
            self.root.protocol("WM_DELETE_WINDOW", self._on_quit)

    # ----- UI 構築 -----
    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        frame_main = ttk.Frame(nb)
        nb.add(frame_main, text="  📊 監視  ")
        self._build_main_tab(frame_main)

        frame_cfg = ttk.Frame(nb)
        nb.add(frame_cfg, text="  ⚙ 設定  ")
        self._build_settings_tab(frame_cfg)

        frame_map = ttk.Frame(nb)
        nb.add(frame_map, text="  🗺 マッピング  ")
        self._build_mapping_tab(frame_map)

    def _build_main_tab(self, parent):
        ctrl = ttk.Frame(parent)
        ctrl.pack(fill=tk.X, padx=8, pady=6)

        self.btn_start = ttk.Button(ctrl, text="▶ 監視開始",
            command=self.start_monitoring, width=14)
        self.btn_start.pack(side=tk.LEFT, padx=4)

        self.btn_stop = ttk.Button(ctrl, text="■ 停止",
            command=self.stop_monitoring, width=10, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, padx=4)

        self.btn_run_now = ttk.Button(ctrl, text="⚡ 今すぐ実行",
            command=self.run_now, width=12)
        self.btn_run_now.pack(side=tk.LEFT, padx=4)

        self.btn_reload_map = ttk.Button(ctrl, text="🔄 マッピング再読み込み",
            command=self.reload_mapping, width=20)
        self.btn_reload_map.pack(side=tk.LEFT, padx=4)

        status_fr = ttk.LabelFrame(parent, text="状態")
        status_fr.pack(fill=tk.X, padx=8, pady=(0, 4))
        self.lbl_status = ttk.Label(status_fr, text="待機中", foreground="gray")
        self.lbl_status.pack(side=tk.LEFT, padx=8, pady=4)
        self.lbl_stats = ttk.Label(status_fr, text="書込:0 | エラー:0 | スキップ:0",
                                   foreground="gray")
        self.lbl_stats.pack(side=tk.RIGHT, padx=8, pady=4)

        log_fr = ttk.LabelFrame(parent, text="ログ")
        log_fr.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))

        self.log_text = tk.Text(log_fr, state=tk.DISABLED, wrap=tk.WORD,
                                font=("Consolas", 10), bg="#1e1e1e", fg="#d4d4d4",
                                insertbackground="white", relief=tk.FLAT)
        scroll = ttk.Scrollbar(log_fr, orient=tk.VERTICAL,
                               command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        btn_fr = ttk.Frame(parent)
        btn_fr.pack(fill=tk.X, padx=8, pady=(0, 4))
        ttk.Button(btn_fr, text="ログをクリア",
                   command=self._clear_log).pack(side=tk.RIGHT, padx=4)

    def _build_settings_tab(self, parent):
        fr = ttk.Frame(parent)
        fr.pack(fill=tk.BOTH, padx=16, pady=12)

        fields = [
            ("Firebase URL", "firebase_url", False),
            ("Web API キー", "firebase_api_key", False),
            ("メールアドレス", "firebase_email", False),
            ("パスワード", "firebase_password", True),
        ]
        self._cfg_vars = {}
        for label, key, secret in fields:
            row = ttk.Frame(fr)
            row.pack(fill=tk.X, pady=4)
            ttk.Label(row, text=label, width=18, anchor=tk.W).pack(side=tk.LEFT)
            var = tk.StringVar(value=self.cfg.get(key, ""))
            self._cfg_vars[key] = var
            show = "*" if secret else ""
            entry = ttk.Entry(row, textvariable=var, show=show, width=48)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        row = ttk.Frame(fr)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text="Excel フォルダ", width=18, anchor=tk.W).pack(side=tk.LEFT)
        var = tk.StringVar(value=self.cfg.get("excel_folder", ""))
        self._cfg_vars["excel_folder"] = var
        ttk.Entry(row, textvariable=var, width=38).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="参照...",
            command=lambda: var.set(
                filedialog.askdirectory(title="Excelフォルダを選択") or var.get()
            )).pack(side=tk.LEFT, padx=4)

        auto_row = ttk.Frame(fr)
        auto_row.pack(fill=tk.X, pady=6)
        self._auto_var = tk.BooleanVar(value=self.cfg.get("auto_start", True))
        ttk.Checkbutton(auto_row, text="起動時に自動で監視を開始する",
                        variable=self._auto_var).pack(side=tk.LEFT)

        btn_row = ttk.Frame(fr)
        btn_row.pack(fill=tk.X, pady=8)
        ttk.Button(btn_row, text="💾 設定を保存",
                   command=self._save_settings, width=16).pack(side=tk.LEFT, padx=4)
        ttk.Button(btn_row, text="🔗 接続テスト",
                   command=self._test_connection, width=14).pack(side=tk.LEFT, padx=4)

        self.lbl_test_result = ttk.Label(fr, text="", foreground="gray")
        self.lbl_test_result.pack(fill=tk.X, pady=4)

    def _build_mapping_tab(self, parent):
        fr = ttk.Frame(parent)
        fr.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        ttk.Label(fr, text="pool_mapping.json の内容", font=("", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(fr, text=f"ファイル: {MAPPING_FILE}",
                  foreground="gray", font=("", 9)).pack(anchor=tk.W)

        self.map_text = tk.Text(fr, state=tk.DISABLED, wrap=tk.NONE,
                                font=("Consolas", 9))
        scroll_y = ttk.Scrollbar(fr, orient=tk.VERTICAL, command=self.map_text.yview)
        scroll_x = ttk.Scrollbar(fr, orient=tk.HORIZONTAL, command=self.map_text.xview)
        self.map_text.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.map_text.pack(fill=tk.BOTH, expand=True)

        ttk.Button(fr, text="🔄 再読み込み",
                   command=self.reload_mapping).pack(anchor=tk.W, pady=4)

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _update_status_label(self):
        if self.worker and self.worker.is_running:
            if _sse_connected.is_set():
                status = "🟢 リアルタイム監視中"
            else:
                status = "🔵 ポーリング監視中"
        elif self.poller and self.poller.is_running:
            status = "🔵 ポーリング監視中"
        else:
            status = "待機中"
        self.lbl_status.config(text=status)

        if self.worker:
            stats = self.worker.get_stats()
            self.lbl_stats.config(
                text=f"書込:{stats['write']} | エラー:{stats['error']} | スキップ:{stats['skip']}")

    def start_monitoring(self):
        if requests is None:
            messagebox.showerror("エラー",
                "requests ライブラリがインストールされていません。\n"
                "pip install requests を実行してください。")
            return

        self.cfg = load_config()
        _update_firebase_credentials(self.cfg)

        self.worker = PoolWriteWorker(self.cfg, self.mapping, self.log)
        self.worker.start()

        self.poller = PoolPoller(self.cfg, self.log)
        self.poller.start()

        if SSE_AVAILABLE:
            self.sse_listener = PoolSSEListener(self.cfg, self.log, self._on_sse_status)
            self.sse_listener.start()

        self.btn_start.config(state=tk.DISABLED)
        self.btn_stop.config(state=tk.NORMAL)
        self.log("監視を開始しました")
        self._update_status_label()
        self.root.after(1000, self._update_ui)

    def stop_monitoring(self):
        if self.worker:
            self.worker.stop()
            self.worker = None
        if self.poller:
            self.poller.stop()
            self.poller = None
        if self.sse_listener:
            self.sse_listener.stop()
            self.sse_listener = None

        self.btn_start.config(state=tk.NORMAL)
        self.btn_stop.config(state=tk.DISABLED)
        self.log("監視を停止しました")
        self._update_status_label()

    def run_now(self):
        if self.worker and self.worker.is_running:
            records = firebase_get_pending(self.cfg.get("firebase_url", ""),
                                          self.cfg.get("firebase_api_key", ""))
            if records:
                for rid in records:
                    enqueue_record(rid)
                self.log(f"[Manual] {len(records)} 件をキューに投入")
            else:
                self.log("[Manual] 未処理レコード: なし")
        else:
            messagebox.showwarning("警告", "監視が実行中ではありません")

    def reload_mapping(self):
        self.mapping = load_mapping()
        if self.worker:
            self.worker.reload_mapping(self.mapping)
        self.log(f"マッピングを再読み込みしました ({len(self.mapping)} 校)")

        self.map_text.config(state=tk.NORMAL)
        self.map_text.delete(1.0, tk.END)
        self.map_text.insert(tk.END, json.dumps(self.mapping, ensure_ascii=False, indent=2))
        self.map_text.config(state=tk.DISABLED)

    def _on_sse_status(self, state, msg):
        if state == "connected":
            self.log(f"[SSE] {msg}")
            _sse_connected.set()
        elif state == "reconnecting":
            self.log(f"[SSE] {msg}")
        elif state == "stopped":
            _sse_connected.clear()
        self._update_status_label()

    def _save_settings(self):
        for key, var in self._cfg_vars.items():
            self.cfg[key] = var.get()
        self.cfg["auto_start"] = self._auto_var.get()
        save_config(self.cfg)
        _update_firebase_credentials(self.cfg)
        self.log("設定を保存しました")
        self.lbl_test_result.config(text="")

    def _test_connection(self):
        firebase_url = self._cfg_vars["firebase_url"].get().strip()
        api_key = self._cfg_vars["firebase_api_key"].get().strip()

        if not firebase_url or not api_key:
            self.lbl_test_result.config(text="⚠ Firebase URL と Web API キーを入力してください", foreground="red")
            return

        self.lbl_test_result.config(text="接続テスト中...", foreground="blue")
        self.root.update()

        try:
            token = _get_firebase_token(api_key)
            if token:
                # 接続テストも _http_request 経由で SSL EOF を耐性化
                url = firebase_url.rstrip("/") + "/pool_inspections.json"
                resp = _http_request("GET", url, api_key=api_key, timeout=10,
                                     max_attempts=2, log_prefix="TEST")
                if resp is not None and resp.status_code == 200:
                    self.lbl_test_result.config(text="✓ 接続成功", foreground="green")
                elif resp is not None:
                    self.lbl_test_result.config(text=f"✗ HTTP {resp.status_code}", foreground="red")
                else:
                    self.lbl_test_result.config(text="✗ 接続失敗（リトライ後も到達不能）", foreground="red")
            else:
                self.lbl_test_result.config(text="✗ トークン取得失敗", foreground="red")
        except Exception as e:
            self.lbl_test_result.config(text=f"✗ エラー: {e}", foreground="red")

    def _update_ui(self):
        if self.worker and self.worker.is_running:
            self._update_status_label()
            self.root.after(1000, self._update_ui)

    def _on_close_to_tray(self):
        self.root.withdraw()
        if not self._tray_icon and TRAY_AVAILABLE:
            self._create_tray_icon()

    def _create_tray_icon(self):
        if not TRAY_AVAILABLE:
            return
        try:
            img = Image.new('RGB', (64, 64), color='white')
            draw = ImageDraw.Draw(img)
            draw.rectangle([8, 8, 56, 56], outline='blue', width=3)

            menu = pystray.Menu(
                pystray.MenuItem("表示", self._show_window, default=True),
                pystray.MenuItem("終了", self._on_quit),
            )
            self._tray_icon = pystray.Icon(APP_NAME, img, APP_NAME, menu=menu)
            t = threading.Thread(target=self._tray_icon.run, daemon=True)
            t.start()
        except Exception:
            pass

    def _show_window(self, icon=None, item=None):
        self.root.after(0, self._do_show_window)

    def _do_show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _on_quit(self, icon=None, item=None):
        self.stop_monitoring()
        if self._tray_icon:
            self._tray_icon.stop()
        self.root.destroy()


def main():
    root = tk.Tk()
    app = PoolWriterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
