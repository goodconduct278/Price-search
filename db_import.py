# -*- coding: utf-8 -*-
"""
CSV → SQLite 統合インポートスクリプト

csv/ フォルダ内の *.csv を全スキャンし、名称・価格を一括取込む。
ファイル名（拡張子なし）= source_db 名。

取込先テーブル:
  product_price_master  品番CD → 正式品名（UPSERT）
  name_aliases          品番CD → 別名・正規化名（UPSERT）
  product_prices        品番CD → 価格（source_db 単位で洗い替え）

CSVの必須列: 品番CD（または 品番コード）、材料名（または 品番名）
オプション列: 価格、荷姿 / 単位、備考
価格列がない場合は product_prices への取込みをスキップする。

使い方:
  python db_import.py --db 商品価格管理.db --csvdir csv
  python db_import.py --db 商品価格管理.db --csvdir csv --source 標準DB
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import unicodedata
from pathlib import Path


# ──────────────────────────────────────────
# 列名マッピング（表記ゆれ対応）
# ──────────────────────────────────────────

COL_PRODUCT_CD = ["品番CD", "品番コード", "品番cd"]
COL_NAME       = ["材料名", "品番名", "品名"]
COL_PRICE      = ["価格", "単価", "仕入価格"]
COL_NOTE       = ["備  考", "備考", "備　考"]
COL_UNIT       = ["荷姿 / 単位", "荷姿/単位", "荷姿・単位", "荷姿／単位", "荷姿・寸法"]
# 面積別の特価（任意列）。あれば price_tier1 / price_tier2 に取込む
COL_TIER1      = ["特価1", "特化1", "特価①", "特価➀", "特価1価格"]
COL_TIER2      = ["特価2", "特化2", "特価②", "特価➁", "特価2価格"]

# 面積閾値の設定ファイル（csv/ に置く別ファイルで管理）
# ファイル名（拡張子なし）がこれらのいずれかなら、価格表ではなく閾値定義として読む
SETTINGS_STEMS    = {"面積設定", "面積閾値", "面積しきい値", "エリア設定", "area_settings"}
COL_SETTINGS_DB   = ["DB名", "DB", "価格表", "価格表DB", "ソースDB", "source_db"]
COL_SETTINGS_T1   = ["特価1面積", "特価1の面積", "特価①面積", "特価1", "特化1", "tier1", "tier1_area"]
COL_SETTINGS_T2   = ["特価2面積", "特価2の面積", "特価②面積", "特価2", "特化2", "tier2", "tier2_area"]


def find_col(headers: list[str], candidates: list[str]) -> str | None:
    for c in candidates:
        if c in headers:
            return c
    return None


# ──────────────────────────────────────────
# テキスト正規化
# ──────────────────────────────────────────

def clean_text(value) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if s in ("nan", "None", "NaN"):
        return ""
    return s.replace("\ufeff", "").strip()


def normalize_name(value: str) -> str:
    s = clean_text(value)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("Φ", "φ")
    s = s.replace(" ", "").replace("\u3000", "")
    return s.upper()


def normalize_cd(value: str) -> str:
    s = clean_text(value)
    if not s:
        return ""
    try:
        return str(int(float(s)))
    except ValueError:
        return s


def parse_area(value) -> float | None:
    """'301㎡' / '301 m2' / '1,200' などを数値(㎡)に変換する。空・不正は None"""
    s = clean_text(value)
    if not s:
        return None
    s = unicodedata.normalize("NFKC", s)
    for token in ("㎡", "m2", "M2", "平米", "平方メートル", ",", " "):
        s = s.replace(token, "")
    try:
        return float(s)
    except ValueError:
        return None


# ──────────────────────────────────────────
# DB操作
# ──────────────────────────────────────────

def ensure_tables(conn: sqlite3.Connection):
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS product_price_master (
            product_cd    TEXT PRIMARY KEY,
            official_name TEXT
        );

        CREATE TABLE IF NOT EXISTS name_aliases (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            product_cd      TEXT NOT NULL,
            official_name   TEXT,
            alias_name      TEXT NOT NULL,
            normalized_name TEXT NOT NULL,
            UNIQUE (product_cd, alias_name)
        );
        CREATE INDEX IF NOT EXISTS idx_aliases_norm
            ON name_aliases (normalized_name);

        CREATE TABLE IF NOT EXISTS product_prices (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            product_cd    TEXT NOT NULL,
            material_name TEXT,
            source_db     TEXT NOT NULL,
            price         TEXT,
            unit          TEXT,
            note          TEXT,
            price_tier1   TEXT,
            price_tier2   TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_prices_cd_source
            ON product_prices (product_cd, source_db);

        -- 面積別特価の閾値（DB=価格表ごとに設定。未登録のDBは面積処理なし）
        CREATE TABLE IF NOT EXISTS area_thresholds (
            source_db  TEXT PRIMARY KEY,
            tier1_area REAL,
            tier2_area REAL
        );
    """)
    # 既存DBに列がなければ追加（後方互換マイグレーション）
    for col in ("unit", "note", "price_tier1", "price_tier2"):
        try:
            conn.execute(f"ALTER TABLE product_prices ADD COLUMN {col} TEXT")
        except sqlite3.OperationalError:
            pass  # すでに存在する
    conn.commit()


def import_one_csv(conn: sqlite3.Connection, csv_path: Path, source_db: str) -> dict:
    rows = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []

        col_cd    = find_col(headers, COL_PRODUCT_CD)
        col_name  = find_col(headers, COL_NAME)
        col_price = find_col(headers, COL_PRICE)
        col_note  = find_col(headers, COL_NOTE)
        col_unit  = find_col(headers, COL_UNIT)
        col_tier1 = find_col(headers, COL_TIER1)
        col_tier2 = find_col(headers, COL_TIER2)

        if not col_cd:
            raise ValueError(f"品番CD列が見つかりません: {headers}")
        if not col_name:
            raise ValueError(f"材料名列が見つかりません: {headers}")

        has_price = col_price is not None

        for row in reader:
            cd   = normalize_cd(row.get(col_cd, ""))
            name = clean_text(row.get(col_name, ""))
            if not cd or not name:
                continue

            price = clean_text(row.get(col_price, "")) if has_price else ""
            unit  = clean_text(row.get(col_unit, ""))  if col_unit  else ""
            note  = clean_text(row.get(col_note,  ""))  if col_note  else ""
            tier1 = clean_text(row.get(col_tier1, "")) if col_tier1 else ""
            tier2 = clean_text(row.get(col_tier2, "")) if col_tier2 else ""
            rows.append({"cd": cd, "name": name, "price": price, "unit": unit,
                         "note": note, "tier1": tier1, "tier2": tier2})

    if not rows:
        print(f"  [{source_db}] {csv_path.name}: 有効行なし、スキップ")
        return {"master": 0, "alias": 0, "price": 0}

    cnt = {"master": 0, "alias": 0, "price": 0}

    with conn:
        # product_price_master: UPSERT（既存は上書き）
        conn.executemany(
            "INSERT OR REPLACE INTO product_price_master (product_cd, official_name) VALUES (?, ?)",
            [(r["cd"], r["name"]) for r in rows]
        )
        cnt["master"] = len(rows)

        # name_aliases: UPSERT（既存は normalized_name を更新）
        alias_rows = [
            (r["cd"], r["name"], r["name"], normalize_name(r["name"]))
            for r in rows
        ]
        conn.executemany("""
            INSERT INTO name_aliases (product_cd, official_name, alias_name, normalized_name)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(product_cd, alias_name)
            DO UPDATE SET
                official_name   = excluded.official_name,
                normalized_name = excluded.normalized_name
        """, alias_rows)
        cnt["alias"] = len(rows)

        # product_prices: source_db 単位で洗い替え（価格列がある場合のみ）
        if has_price:
            conn.execute(
                "DELETE FROM product_prices WHERE source_db = ?", (source_db,)
            )
            conn.executemany(
                "INSERT INTO product_prices "
                "(product_cd, material_name, source_db, price, unit, note, price_tier1, price_tier2) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                [(r["cd"], r["name"], source_db, r["price"], r["unit"], r["note"],
                  r["tier1"], r["tier2"]) for r in rows]
            )
            cnt["price"] = len(rows)

    label = f"master:{cnt['master']} alias:{cnt['alias']} price:{cnt['price'] if has_price else 'なし（価格列なし）'}"
    print(f"  [{source_db}] {csv_path.name}: {len(rows)} 行 / {label}")
    return cnt


def load_area_settings(conn: sqlite3.Connection, csv_path: Path) -> int:
    """面積閾値の設定ファイルを読み、area_thresholds に登録する。
    形式: DB名 / 特価1面積 / 特価2面積（列名は表記ゆれ可）"""
    settings = []

    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []

        col_db = find_col(headers, COL_SETTINGS_DB)
        col_t1 = find_col(headers, COL_SETTINGS_T1)
        col_t2 = find_col(headers, COL_SETTINGS_T2)

        if not col_db:
            raise ValueError(f"DB名列が見つかりません: {headers}")

        for row in reader:
            db = clean_text(row.get(col_db, ""))
            if not db:
                continue
            t1 = parse_area(row.get(col_t1, "")) if col_t1 else None
            t2 = parse_area(row.get(col_t2, "")) if col_t2 else None
            # 設定ミス検知: 特価2の面積が特価1の面積以下だと段階が逆転し誤った価格になる
            if t1 is not None and t2 is not None and t2 <= t1:
                print(f"  [面積設定] 警告: {db} の特価2面積({t2:g})が特価1面積({t1:g})以下です。"
                      f"特価1<特価2 の順に設定してください。")
            settings.append((db, t1, t2))

    with conn:
        for db, t1, t2 in settings:
            conn.execute("""
                INSERT INTO area_thresholds (source_db, tier1_area, tier2_area)
                VALUES (?, ?, ?)
                ON CONFLICT(source_db) DO UPDATE SET
                    tier1_area = excluded.tier1_area,
                    tier2_area = excluded.tier2_area
            """, (db, t1, t2))

    print(f"  [面積設定] {csv_path.name}: {len(settings)} 件のDB閾値を登録")
    return len(settings)


# ──────────────────────────────────────────
# メイン
# ──────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description="CSVをSQLiteにインポートする（名称・価格統合）")
    parser.add_argument("--db",     required=True, help="SQLiteDBファイルパス")
    parser.add_argument("--csvdir", required=True, help="CSVフォルダ")
    parser.add_argument("--source", default=None,  help="特定ソースのみ処理（例: 標準DB）")
    args = parser.parse_args()

    db_path  = Path(args.db)
    csv_dir  = Path(args.csvdir)

    if not csv_dir.exists():
        raise SystemExit(f"CSVフォルダが見つかりません: {csv_dir}")

    csv_files = sorted(csv_dir.glob("*.csv"))
    if not csv_files:
        raise SystemExit(f"CSVファイルが見つかりません: {csv_dir}")

    if args.source:
        # 面積設定ファイルは価格表ではないので --source 指定時も常に処理する
        # （閾値の編集が特定ソース取込でも反映されるように）
        target = [f for f in csv_files if f.stem == args.source]
        if not target:
            raise SystemExit(f"指定ソースのCSVが見つかりません: {args.source}")
        settings = [f for f in csv_files if f.stem in SETTINGS_STEMS and f not in target]
        csv_files = target + settings

    conn = sqlite3.connect(db_path)
    try:
        ensure_tables(conn)

        total_rows = 0
        errors = []

        for csv_path in csv_files:
            source_db = csv_path.stem

            # 面積閾値の設定ファイルは価格表ではなく閾値定義として読む
            if source_db in SETTINGS_STEMS:
                try:
                    load_area_settings(conn, csv_path)
                except Exception as e:
                    errors.append((source_db, str(e)))
                    print(f"  [{source_db}] ERROR: {e}")
                continue

            try:
                cnt = import_one_csv(conn, csv_path, source_db)
                total_rows += cnt["master"]
            except Exception as e:
                errors.append((source_db, str(e)))
                print(f"  [{source_db}] ERROR: {e}")

        print(f"\n合計 {total_rows} 件インポート。DB: {db_path}")
        if errors:
            print(f"エラー {len(errors)} 件:")
            for src, msg in errors:
                print(f"  - [{src}] {msg}")
            return 1

    finally:
        conn.close()

    return 0


if __name__ == "__main__":
    import sys
    try:
        raise SystemExit(main())
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"ERROR: {e}", file=sys.stderr)
        raise SystemExit(1)
