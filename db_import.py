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
            note          TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_prices_cd_source
            ON product_prices (product_cd, source_db);
    """)
    # 既存DBにnote列がなければ追加
    try:
        conn.execute("ALTER TABLE product_prices ADD COLUMN note TEXT")
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
            rows.append({"cd": cd, "name": name, "price": price, "unit": unit, "note": note})

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
                "INSERT INTO product_prices (product_cd, material_name, source_db, price, unit, note) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                [(r["cd"], r["name"], source_db, r["price"], r["unit"], r["note"]) for r in rows]
            )
            cnt["price"] = len(rows)

    label = f"master:{cnt['master']} alias:{cnt['alias']} price:{cnt['price'] if has_price else 'なし（価格列なし）'}"
    print(f"  [{source_db}] {csv_path.name}: {len(rows)} 行 / {label}")
    return cnt


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
        csv_files = [f for f in csv_files if f.stem == args.source]
        if not csv_files:
            raise SystemExit(f"指定ソースのCSVが見つかりません: {args.source}")

    conn = sqlite3.connect(db_path)
    try:
        ensure_tables(conn)

        total_rows = 0
        errors = []

        for csv_path in csv_files:
            source_db = csv_path.stem
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
