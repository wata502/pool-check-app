"""
setup_mapping.py
プール付属電気設備点検システム - マッピング自動生成スクリプト

対象フォルダ内の xlsm ファイルを自動スキャンし、
pool_mapping.json を生成する。初回のみ実行する。

使い方:
  python setup_mapping.py [Excel格納フォルダのパス]
  ※ 引数省略時は input() でパスを入力
"""

from __future__ import annotations  # SM-1: Python 3.9以前との互換性

import os
import sys
import json
import re
import traceback
from pathlib import Path

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("[エラー] pywin32 がインストールされていません。")
    print("  pip install pywin32 を実行してから再試行してください。")
    sys.exit(1)

# ========== 設定 ==========

# スクリプトと同じフォルダに保存
OUTPUT_FILE = Path(__file__).parent / "pool_mapping.json"

# ファイル番号 → 学校名（ファイル名から自動抽出するが念のため）
FILE_NUMBER_RE = re.compile(r'^(\d{3})')

# 固定行範囲（全ファイル共通）
INSULATION_ROWS  = list(range(16, 23))   # R16-R22（最大7行）
GROUNDING_ROWS   = list(range(26, 30))   # R26-R29（最大4行）
BREAKER_ROWS     = list(range(33, 38))   # R33-R37（最大5行）

GROUNDING_METHOD_CELL = "H23"

# セクション別書き込みカラム
INSULATION_COLS  = {"value": "D", "judgment": "F", "remarks": "G"}
GROUNDING_COLS   = {"value": "D", "judgment": "F", "remarks": "G"}
BREAKER_TEST_COLS = {"judgment": "F", "remarks": "G"}

# ヘッダーセル
HEADER_MAP = {
    "date_cell":        "B5",
    "time_cell":        "B6",
    "weather_cell":     "D6",
    "temperature_cell": "B7",
    "humidity_cell":    "D7",
    "note_cell":        "E5",
}

SPECIAL_NOTES_RANGE = {"start": "A39", "end": "H49"}

# ========== ユーティリティ ==========

def col_to_num(col_str: str) -> int:
    """列記号 → 1始まり列番号"""
    result = 0
    for c in col_str.upper():
        result = result * 26 + (ord(c) - ord('A') + 1)
    return result

def num_to_col(n: int) -> str:
    """1始まり列番号 → 列記号"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord('A') + r) + s
    return s

def safe_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def find_grounding_method_cell(ws, default_cell: str = "H23"):
    """
    接地抵抗測定セクション周辺で「測定方式」ラベルを動的検出し、
    その右隣のセル位置(例: "H24")と現在値を返す。

    走査範囲: 行20〜25, 列A〜G（右隣がH列以内に収まるよう G列まで）
    見つからなければ default_cell をそのまま返す（後方互換）。

    ※ シートごとに1〜2行のレイアウトずれ（H23/H24混在）を吸収するための仕組み。
       検出範囲外まで広げると誤検出のリスクが上がるため、意図的に狭く制限。
    """
    try:
        for r in range(20, 26):
            for c in range(1, 8):  # A〜G
                v = safe_str(ws.Cells(r, c).Value)
                if "測定方式" in v:
                    target_col = c + 1
                    cell_addr = f"{num_to_col(target_col)}{r}"
                    cell_val = safe_str(ws.Cells(r, target_col).Value)
                    return cell_addr, cell_val
    except Exception:
        # スキャン中の例外はフォールバック扱い（堅牢性優先）
        pass

    # フォールバック: default_cell の値を読む
    try:
        mc = re.sub(r'\d', '', default_cell).upper()
        mr = int(re.sub(r'[A-Za-z]', '', default_cell))
        cell_val = safe_str(ws.Cells(mr, col_to_num(mc)).Value)
    except Exception:
        cell_val = ""
    return default_cell, cell_val


# ========== Excel スキャン ==========

def scan_file(xl_app, file_path: str) -> tuple[str, dict] | None:
    """
    1ファイルをスキャンしてマッピング辞書を返す。
    エラー時は None を返す。
    """
    abs_path = os.path.normpath(os.path.abspath(file_path))
    print(f"  スキャン中: {os.path.basename(abs_path)}")

    wb = None
    try:
        wb = xl_app.Workbooks.Open(abs_path, ReadOnly=True, UpdateLinks=False)

        # シート特定（ファイル名の数字部分を除いた名前 or 最初のシート）
        ws = None
        school_name_from_sheet = ""
        for i in range(1, wb.Sheets.Count + 1):
            sname = wb.Sheets(i).Name
            # 数字3桁_で始まらないシートを優先（= 学校名シート）
            if not re.match(r'^\d', sname) and sname not in ("Sheet1", "Sheet2", "Sheet3"):
                ws = wb.Sheets(i)
                school_name_from_sheet = sname
                break
        if ws is None:
            ws = wb.Sheets(1)
            school_name_from_sheet = ws.Name

        # ファイル名から番号を抽出
        base = os.path.basename(abs_path)
        m = FILE_NUMBER_RE.match(base)
        file_id = m.group(1) if m else base.replace(".xlsm", "")

        # 学校名（ファイル名 or シート名から）
        school_name = school_name_from_sheet or base.replace(".xlsm", "")

        # ========== 絶縁抵抗測定 スキャン ==========
        insulation = []
        current_panel = ""
        for row in INSULATION_ROWS:
            panel   = safe_str(ws.Cells(row, col_to_num("A")).Value)
            circuit = safe_str(ws.Cells(row, col_to_num("B")).Value)
            voltage_raw = ws.Cells(row, col_to_num("C")).Value
            try:
                voltage = int(float(str(voltage_raw))) if voltage_raw else None
            except (ValueError, TypeError):
                voltage = None

            # フォワードフィル: A列結合による空欄を直前の値で補完
            if panel:
                current_panel = panel
            else:
                panel = current_panel

            # 固定行範囲を全走査。空白行(分電盤区切り)があっても break しない。
            # 測定回路(B列)が空ならその行はマッピング対象外として continue。
            if not circuit:
                continue

            insulation.append({
                "row": row,
                "panel": panel,
                "circuit": circuit,
                "voltage": voltage,
            })

        # ========== 接地抵抗測定 スキャン ==========
        grounding = []
        current_panel = ""
        for row in GROUNDING_ROWS:
            panel      = safe_str(ws.Cells(row, col_to_num("A")).Value)
            circuit    = safe_str(ws.Cells(row, col_to_num("B")).Value)
            grnd_type  = safe_str(ws.Cells(row, col_to_num("C")).Value)

            # フォワードフィル: A列結合による空欄を直前の値で補完
            if panel:
                current_panel = panel
            else:
                panel = current_panel

            # 全走査・break廃止。circuit が空の行はマッピング対象外。
            if not circuit:
                continue

            grounding.append({
                "row": row,
                "panel": panel,
                "circuit": circuit,
                "type": grnd_type,
            })

        # 測定方式: 「測定方式」ラベルを動的検出して右隣のセル位置と値を採用
        # （シート毎に H23 / H24 等の差異を吸収。検出失敗時は GROUNDING_METHOD_CELL にフォールバック）
        grounding_method_cell_addr, grounding_method_val = \
            find_grounding_method_cell(ws, GROUNDING_METHOD_CELL)

        # ========== 漏電遮断器テスト スキャン ==========
        breaker_test = []
        current_panel = ""
        for row in BREAKER_ROWS:
            panel   = safe_str(ws.Cells(row, col_to_num("A")).Value)
            circuit = safe_str(ws.Cells(row, col_to_num("B")).Value)
            cap_raw = ws.Cells(row, col_to_num("C")).Value
            try:
                capacity = int(float(str(cap_raw))) if cap_raw else None
            except (ValueError, TypeError):
                capacity = None

            # フォワードフィル: A列結合による空欄を直前の値で補完
            if panel:
                current_panel = panel
            else:
                panel = current_panel

            # 全走査・break廃止。circuit が空の行はマッピング対象外。
            if not circuit:
                continue

            breaker_test.append({
                "row": row,
                "panel": panel,
                "circuit": circuit,
                "capacity": capacity,
            })

        # has_data: 少なくとも1つのセクションにデータがあれば True
        has_data = bool(insulation or grounding or breaker_test)

        mapping = {
            "file_name": base,
            "school_name": school_name,
            "sheet_name": school_name_from_sheet,
            "has_data": has_data,
            "insulation": insulation,
            "insulation_cols": INSULATION_COLS,
            "grounding": grounding,
            "grounding_cols": GROUNDING_COLS,
            "grounding_method_cell": grounding_method_cell_addr,
            "grounding_method": grounding_method_val,
            "breaker_test": breaker_test,
            "breaker_test_cols": BREAKER_TEST_COLS,
            "header": HEADER_MAP,
            "special_notes_range": SPECIAL_NOTES_RANGE,
        }

        return file_id, mapping

    except Exception as e:
        print(f"    [エラー] {os.path.basename(file_path)}: {e}")
        traceback.print_exc()
        return None
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass


def scan_folder(folder_path: str) -> dict:
    """フォルダ内の全 xlsm をスキャンしてマッピング辞書を返す"""
    folder = Path(folder_path)
    if not folder.is_dir():
        print(f"[エラー] フォルダが見つかりません: {folder_path}")
        return {}

    xlsm_files = sorted(folder.rglob("*.xlsm"))
    if not xlsm_files:
        xlsm_files = sorted(folder.rglob("*.xlsx"))

    print(f"\n対象ファイル: {len(xlsm_files)} 件")
    if not xlsm_files:
        print("  xlsm/xlsx ファイルが見つかりません。")
        return {}

    pythoncom.CoInitialize()
    xl_app = None
    mapping_all = {}

    try:
        xl_app = win32com.client.DispatchEx("Excel.Application")
        xl_app.Visible = False
        xl_app.DisplayAlerts = False
        xl_app.ScreenUpdating = False
        xl_app.EnableEvents = False

        for fpath in xlsm_files:
            result = scan_file(xl_app, str(fpath))
            if result is None:
                continue
            file_id, mapping = result
            mapping_all[file_id] = mapping
            status = "✓ データあり" if mapping["has_data"] else "  (データなし)"
            print(f"    {status}  絶縁:{len(mapping['insulation'])}行  "
                  f"接地:{len(mapping['grounding'])}行  "
                  f"漏電:{len(mapping['breaker_test'])}行  "
                  f"({mapping['school_name']})")

    except Exception as e:
        print(f"[致命的エラー] {e}")
        traceback.print_exc()
    finally:
        if xl_app is not None:
            try:
                xl_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    return mapping_all


def save_mapping(mapping_all: dict, output_path: str):
    """マッピングをJSONに保存する"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(mapping_all, f, ensure_ascii=False, indent=2)
    print(f"\n[保存] {output_path}")
    print(f"  学校数: {len(mapping_all)}")


# ========== エントリーポイント ==========

def main():
    if len(sys.argv) >= 2:
        folder = sys.argv[1]
    else:
        folder = input("Excel格納フォルダのパスを入力してください: ").strip().strip('"')

    if not folder or not os.path.isdir(folder):
        print(f"[エラー] フォルダが存在しません: {folder}")
        sys.exit(1)

    mapping_all = scan_folder(folder)
    if mapping_all:
        save_mapping(mapping_all, str(OUTPUT_FILE))
        print("\n[完了] pool_mapping.json を生成しました。")
    else:
        print("\n[警告] 有効なマッピングが見つかりませんでした。")


if __name__ == "__main__":
    main()
