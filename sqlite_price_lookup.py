# -*- coding: utf-8 -*-
"""
SQLite価格検索 Python側スクリプト
ポップアップ候補選択対応版 ＋ 荷姿/単位（unit）出力対応版

動き:
1. 完全一致が1件なら自動で価格取得
2. 完全一致なし、または完全一致が複数なら候補を返す
3. Excel VBA側でポップアップ選択する
4. 価格は「B2で選んだDB → 標準DB → 見積DB」の順でフォールバックする

変更点（荷姿/単位対応）:
- get_price_for_source: SELECT に unit を追加し、返り値に含める
- get_best_price: unit を持ち回す（価格未入力でも荷姿は拾えれば返す）
- auto_result_from_cd / 各 result: unit を含める
- enrich_candidates_with_price / pack_candidates: 候補にも unit を含める（末尾に追加）
- write_output_tsv: ヘッダーに unit を追加（official_name と candidates の間）
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import sys
import traceback
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path


REC_SEP = "|||REC|||"
FIELD_SEP = "|||FLD|||"

# DBの能力（列・テーブルの有無）。main() で1回だけ判定して使い回す。
# 旧DB（面積対応前にインポートされたDB）でも落ちないようにするため。
CAPS = {"tier_cols": False, "threshold_tbl": False, "unit_col": False}
# source_db -> (tier1_area, tier2_area) のキャッシュ
_THRESHOLD_CACHE: dict = {}


def parse_area(value) -> float | None:
    """'301㎡' / '301 m2' / '1,200' などを数値(㎡)に変換する。空・不正は None"""
    if value is None:
        return None
    s = str(value).replace("﻿", "").strip()
    if not s:
        return None
    s = unicodedata.normalize("NFKC", s)
    for token in ("㎡", "m2", "M2", "平米", "平方メートル", ",", " "):
        s = s.replace(token, "")
    try:
        return float(s)
    except ValueError:
        return None


def detect_caps(conn: sqlite3.Connection) -> None:
    """product_prices に特価列・unit列があるか、area_thresholds テーブルがあるかを判定"""
    cols = {row[1] for row in conn.execute("PRAGMA table_info(product_prices)")}
    CAPS["tier_cols"] = "price_tier1" in cols and "price_tier2" in cols
    CAPS["unit_col"] = "unit" in cols
    tbl = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name='area_thresholds'"
    ).fetchone()
    CAPS["threshold_tbl"] = tbl is not None


def get_area_thresholds(conn: sqlite3.Connection, source_db: str):
    """指定ソースの (tier1_area, tier2_area) を返す。未設定なら (None, None)"""
    if not CAPS["threshold_tbl"]:
        return (None, None)
    if source_db in _THRESHOLD_CACHE:
        return _THRESHOLD_CACHE[source_db]
    row = conn.execute(
        "SELECT tier1_area, tier2_area FROM area_thresholds WHERE source_db = ?",
        (source_db,)
    ).fetchone()
    result = (row[0], row[1]) if row is not None else (None, None)
    _THRESHOLD_CACHE[source_db] = result
    return result


def resolve_tier_price(base, tier1, tier2, area, t1_area, t2_area):
    """面積に応じて適用価格とラベルを返す。
    仕様: 面積<t1_area=通常 / t1_area以上=特価1 / t2_area以上=特価2
    ラベル "" は面積処理なし（閾値未設定 or 面積未入力）を意味する。"""
    base = "" if base is None else str(base).strip()
    tier1 = "" if tier1 is None else str(tier1).strip()
    tier2 = "" if tier2 is None else str(tier2).strip()

    # 面積未入力、またはこのDBに閾値設定なし → 通常価格（面積処理なし）
    if area is None or (t1_area is None and t2_area is None):
        return base, ""

    if t2_area is not None and area >= t2_area and tier2 != "":
        return tier2, "特価2"
    if t1_area is not None and area >= t1_area and tier1 != "":
        return tier1, "特価1"
    return base, "通常"


def get_valid_sources(conn: sqlite3.Connection) -> list[str]:
    """product_prices に登録済みの source_db 一覧を取得する"""
    rows = conn.execute(
        "SELECT DISTINCT source_db FROM product_prices WHERE source_db IS NOT NULL AND source_db <> '' ORDER BY source_db"
    ).fetchall()
    return [row[0] for row in rows]


def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).replace("\ufeff", "").strip()


def normalize_name(value: str) -> str:
    s = clean_text(value)
    s = unicodedata.normalize("NFKC", s)
    s = s.replace("（", "(").replace("）", ")")
    s = s.replace("Φ", "φ")
    s = s.replace(" ", "").replace("　", "")
    return s.upper()


def read_input_tsv(path: Path):
    rows = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter="\t")
        for row in reader:
            rows.append({
                "row_no": clean_text(row.get("row_no", "")),
                "product_name": clean_text(row.get("product_name", "")),
            })
    return rows


def make_source_order(primary_source: str):
    order = [primary_source, "標準DB", "見積DB"]
    result = []
    for source in order:
        if source and source not in result:
            result.append(source)
    return result


def similarity_score(input_norm: str, alias_norm: str) -> float:
    if not input_norm or not alias_norm:
        return 0.0
    if input_norm == alias_norm:
        return 1.0

    ratio = SequenceMatcher(None, input_norm, alias_norm).ratio()

    if input_norm in alias_norm or alias_norm in input_norm:
        shorter = min(len(input_norm), len(alias_norm))
        longer = max(len(input_norm), len(alias_norm))
        contain_score = 0.70 + 0.30 * (shorter / longer)
        ratio = max(ratio, contain_score)

    return ratio


def find_exact_candidates(conn: sqlite3.Connection, product_name: str):
    normalized = normalize_name(product_name)

    sql = """
    SELECT DISTINCT
        na.product_cd,
        COALESCE(ppm.official_name, na.official_name, na.alias_name) AS official_name,
        na.alias_name
    FROM name_aliases AS na
    LEFT JOIN product_price_master AS ppm
        ON ppm.product_cd = na.product_cd
    WHERE
        na.alias_name = ?
        OR na.normalized_name = ?
    ORDER BY na.product_cd
    """

    cur = conn.execute(sql, (product_name, normalized))
    rows = cur.fetchall()

    by_cd = {}
    for product_cd, official_name, alias_name in rows:
        cd = "" if product_cd is None else str(product_cd)
        if cd == "":
            continue

        by_cd[cd] = {
            "product_cd": cd,
            "official_name": "" if official_name is None else str(official_name),
            "match_name": "" if alias_name is None else str(alias_name),
            "score": 1.0,
            "match_type": "exact",
        }

    return list(by_cd.values())


def find_similar_candidates(conn: sqlite3.Connection, product_name: str, limit: int = 6, threshold: float = 0.52):
    input_norm = normalize_name(product_name)

    sql = """
    SELECT
        na.product_cd,
        COALESCE(ppm.official_name, na.official_name, na.alias_name) AS official_name,
        na.alias_name,
        na.normalized_name
    FROM name_aliases AS na
    LEFT JOIN product_price_master AS ppm
        ON ppm.product_cd = na.product_cd
    WHERE na.product_cd IS NOT NULL
      AND na.product_cd <> ''
    """

    cur = conn.execute(sql)
    best_by_cd = {}

    for product_cd, official_name, alias_name, alias_norm in cur.fetchall():
        cd = "" if product_cd is None else str(product_cd)
        if cd == "":
            continue

        alias_name = "" if alias_name is None else str(alias_name)
        alias_norm = normalize_name(alias_name if alias_norm is None else str(alias_norm))
        score = similarity_score(input_norm, alias_norm)

        if score < threshold:
            continue

        current = best_by_cd.get(cd)
        if current is None or score > current["score"]:
            best_by_cd[cd] = {
                "product_cd": cd,
                "official_name": "" if official_name is None else str(official_name),
                "match_name": alias_name,
                "score": score,
                "match_type": "similar",
            }

    return sorted(best_by_cd.values(), key=lambda x: (-x["score"], x["product_cd"]))[:limit]


def get_price_for_source(conn: sqlite3.Connection, product_cd: str, source_db: str):
    # 変更: unit / 特価1 / 特価2 を SELECT に追加（特価列は無いDBでも動くよう分岐）
    if CAPS["unit_col"]:
        unit_expr = "COALESCE(unit, '') AS unit"
    else:
        unit_expr = "'' AS unit"

    if CAPS["tier_cols"]:
        sql = f"""
        SELECT product_cd, material_name, source_db, price,
               {unit_expr},
               COALESCE(note, '') AS note,
               COALESCE(price_tier1, '') AS price_tier1,
               COALESCE(price_tier2, '') AS price_tier2
        FROM product_prices
        WHERE product_cd = ? AND source_db = ?
        ORDER BY id LIMIT 1
        """
    else:
        sql = f"""
        SELECT product_cd, material_name, source_db, price,
               {unit_expr},
               COALESCE(note, '') AS note,
               '' AS price_tier1,
               '' AS price_tier2
        FROM product_prices
        WHERE product_cd = ? AND source_db = ?
        ORDER BY id LIMIT 1
        """

    row = conn.execute(sql, (product_cd, source_db)).fetchone()
    if row is None:
        return None

    return {
        "product_cd": "" if row[0] is None else str(row[0]),
        "material_name": "" if row[1] is None else str(row[1]),
        "source_db": "" if row[2] is None else str(row[2]),
        "price": "" if row[3] is None else str(row[3]).strip(),
        "unit": "" if row[4] is None else str(row[4]).strip(),
        "note": "" if row[5] is None else str(row[5]).strip(),
        "price_tier1": "" if row[6] is None else str(row[6]).strip(),
        "price_tier2": "" if row[7] is None else str(row[7]).strip(),
    }


def get_best_price(conn: sqlite3.Connection, product_cd: str, primary_source: str, area: float | None = None):
    blank_sources = []
    blank_unit = ""  # 価格未入力でも荷姿だけは拾えたら返す

    for source_db in make_source_order(primary_source):
        hit = get_price_for_source(conn, product_cd, source_db)

        if hit is None:
            continue

        # 面積に応じて適用価格を決定（閾値未設定/面積未入力なら通常価格）
        t1_area, t2_area = get_area_thresholds(conn, source_db)
        price_value, tier_label = resolve_tier_price(
            hit["price"], hit["price_tier1"], hit["price_tier2"],
            area, t1_area, t2_area
        )

        if price_value == "":
            blank_sources.append(source_db)
            if blank_unit == "":
                blank_unit = hit["unit"]
            continue

        return {
            "price_result": price_value,
            "used_source": source_db,
            "price_status": "ok",
            "unit": hit["unit"],
            "note": hit["note"],
            "tier": tier_label,
        }

    if blank_sources:
        return {
            "price_result": "価格未入力",
            "used_source": ",".join(blank_sources),
            "price_status": "blank",
            "unit": blank_unit,
            "note": "",
            "tier": "",
        }

    return {
        "price_result": "価格未入力",
        "used_source": "",
        "price_status": "missing",
        "unit": "",
        "note": "",
        "tier": "",
    }


def enrich_candidates_with_price(conn: sqlite3.Connection, candidates: list[dict], primary_source: str, area: float | None = None):
    enriched = []

    for i, c in enumerate(candidates, start=1):
        price_info = get_best_price(conn, c["product_cd"], primary_source, area)

        enriched.append({
            "no": str(i),
            "product_cd": c["product_cd"],
            "official_name": c["official_name"],
            "match_name": c["match_name"],
            "score": f"{round(c['score'] * 100)}",
            "match_type": c.get("match_type", ""),
            "price_result": price_info["price_result"],
            "used_source": price_info["used_source"],
            "price_status": price_info["price_status"],
            "unit": price_info["unit"],
            "note": price_info["note"],
        })

    return enriched


def safe_candidate_value(value: str) -> str:
    s = clean_text(value)
    s = s.replace(REC_SEP, " ").replace(FIELD_SEP, " ")
    s = s.replace("\t", " ").replace("\r", " ").replace("\n", " ")
    return s


def pack_candidates(candidates: list[dict]) -> str:
    records = []
    for c in candidates:
        # 変更: unit を末尾(fields[8])に追加。既存 fields[0..7] の位置は不変。
        fields = [
            c["no"],
            c["product_cd"],
            c["official_name"],
            c["match_name"],
            c["used_source"],
            c["price_result"],
            c["score"],
            c["note"],
            c["unit"],
        ]
        records.append(FIELD_SEP.join(safe_candidate_value(v) for v in fields))
    return REC_SEP.join(records)


def tier_memo(tier_label: str, area: float | None) -> str:
    """適用された面積区分をメモ用文字列にする（通常/面積処理なしは空）"""
    if tier_label in ("特価1", "特価2"):
        area_txt = "" if area is None else f"・面積{format_area(area)}㎡"
        return f" / {tier_label}適用{area_txt}"
    return ""


def format_area(area: float) -> str:
    """301.0 -> '301' のように見やすく整形"""
    if area == int(area):
        return str(int(area))
    return str(area)


def auto_result_from_cd(conn: sqlite3.Connection, primary_source: str, product_cd: str, official_name: str, match_score: float | None = None, area: float | None = None):
    # match_score を渡すと「類似一致を自動確定した」旨をメモに明示する（完全一致では None）
    sim_note = ""
    if match_score is not None and match_score < 1.0:
        sim_note = f" / 類似度:{round(match_score * 100)}%（自動確定）"

    price_info = get_best_price(conn, product_cd, primary_source, area)

    if price_info["price_status"] == "ok":
        tnote = tier_memo(price_info.get("tier", ""), area)
        if price_info["used_source"] == primary_source:
            memo = f"OK / 品番CD:{product_cd} / 参照DB:{price_info['used_source']}{tnote}{sim_note}"
        else:
            memo = f"代替参照 / 品番CD:{product_cd} / 参照DB:{price_info['used_source']}{tnote}{sim_note}"

        return {
            "price_result": price_info["price_result"],
            "memo": memo,
            "product_cd": product_cd,
            "official_name": official_name,
            "unit": price_info["unit"],
            "candidates": "",
        }

    return {
        "price_result": "価格未入力",
        "memo": f"品番CD:{product_cd} / 対象DBに価格なし{sim_note}",
        "product_cd": product_cd,
        "official_name": official_name,
        "unit": price_info["unit"],
        "candidates": "",
    }


def selection_result(conn: sqlite3.Connection, primary_source: str, candidates: list[dict], memo: str, area: float | None = None):
    enriched = enrich_candidates_with_price(conn, candidates, primary_source, area)

    return {
        "price_result": "候補選択",
        "memo": memo,
        "product_cd": ",".join(c["product_cd"] for c in enriched),
        "official_name": "",
        "unit": "",  # 選択確定後にVBA側で候補から埋める
        "candidates": pack_candidates(enriched),
    }


def get_candidates_in_source(conn: sqlite3.Connection, candidates: list[dict], source_db: str) -> list[dict]:
    """候補リストのうち、指定ソースDBに価格エントリがあるものだけ返す"""
    result = []
    for c in candidates:
        row = conn.execute(
            "SELECT 1 FROM product_prices WHERE product_cd = ? AND source_db = ? LIMIT 1",
            (c["product_cd"], source_db)
        ).fetchone()
        if row is not None:
            result.append(c)
    return result


def lookup_one(conn: sqlite3.Connection, primary_source: str, product_name: str, area: float | None = None):
    exact = find_exact_candidates(conn, product_name)

    if len(exact) > 0:
        # 選択DBを優先順に絞り込む
        # 同一DB内に複数エントリがある場合のみポップアップ
        # 別DBに同名があるだけの場合はフォールバックで自動解決
        for source in make_source_order(primary_source):
            in_source = get_candidates_in_source(conn, exact, source)
            if len(in_source) == 1:
                return auto_result_from_cd(conn, primary_source, in_source[0]["product_cd"], in_source[0]["official_name"], area=area)
            elif len(in_source) >= 2:
                return selection_result(conn, primary_source, in_source, "同じDBに複数の候補があります。候補から選んでください。", area)

        # どのDBにも価格エントリがない場合
        c = exact[0]
        return auto_result_from_cd(conn, primary_source, c["product_cd"], c["official_name"], area=area)

    # 完全一致なし → 類似候補（選択DBで絞り込む）
    similar = find_similar_candidates(conn, product_name)

    if len(similar) == 0:
        return {
            "price_result": "未登録",
            "memo": "名称DBに該当なし",
            "product_cd": "",
            "official_name": "",
            "unit": "",
            "candidates": "",
        }

    # 選択DBの優先順で絞り込み、価格のある候補のみ表示
    # 類似一致は高スコア(>=0.85)のときだけ自動確定し、それ以外は1件でも候補確認を出す
    # （低スコアの名称を無確認で採用すると誤った価格が紛れ込むため）
    AUTO_SIMILAR_THRESHOLD = 0.85
    for source in make_source_order(primary_source):
        in_source = get_candidates_in_source(conn, similar, source)
        if len(in_source) == 1:
            c = in_source[0]
            if c["score"] >= AUTO_SIMILAR_THRESHOLD:
                return auto_result_from_cd(conn, primary_source, c["product_cd"], c["official_name"], match_score=c["score"], area=area)
            return selection_result(conn, primary_source, in_source, "似た候補が1件あります。内容を確認して選んでください。", area)
        elif len(in_source) >= 2:
            return selection_result(conn, primary_source, in_source, "似た候補があります。候補から選んでください。", area)

    return {
        "price_result": "未登録",
        "memo": "選択DBに該当する価格なし",
        "product_cd": "",
        "official_name": "",
        "unit": "",
        "candidates": "",
    }


def write_output_tsv(path: Path, rows):
    # 変更: unit を official_name と candidates の間に追加
    headers = ["row_no", "product_name", "price_result", "memo", "product_cd", "official_name", "unit", "candidates"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, delimiter="\t")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--db", required=True)
    parser.add_argument("--source", required=True)
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--area", default="", help="面積(㎡)。空なら面積処理なし")
    args = parser.parse_args()

    db_path = Path(args.db)
    input_path = Path(args.input)
    output_path = Path(args.output)
    source_db = clean_text(args.source)
    area = parse_area(args.area)

    # source_db の妥当性チェックは DB オープン後に実施

    if not db_path.exists():
        raise SystemExit(f"DBファイルが見つかりません: {db_path}")

    if not input_path.exists():
        raise SystemExit(f"入力TSVが見つかりません: {input_path}")

    input_rows = read_input_tsv(input_path)
    output_rows = []

    conn = sqlite3.connect(db_path)
    try:
        detect_caps(conn)
        valid_sources = get_valid_sources(conn)
        if valid_sources and source_db not in valid_sources:
            raise SystemExit(
                f"検索対象DBが不正です: {source_db}\n"
                f"登録済みのDB: {', '.join(valid_sources)}"
            )

        for row in input_rows:
            product_name = row["product_name"]

            if product_name == "":
                result = {
                    "price_result": "",
                    "memo": "",
                    "product_cd": "",
                    "official_name": "",
                    "unit": "",
                    "candidates": "",
                }
            else:
                result = lookup_one(conn, source_db, product_name, area)

            output_rows.append({
                "row_no": row["row_no"],
                "product_name": product_name,
                **result,
            })
    finally:
        conn.close()

    write_output_tsv(output_path, output_rows)
    return 0


def _write_error_log(message: str) -> None:
    """VBAは非表示ウィンドウでPythonを実行するためstderrが見えない。
    スクリプトと同じフォルダにログを残し、原因を追えるようにする。"""
    try:
        log_path = Path(__file__).with_name("_price_lookup_error.log")
        with log_path.open("w", encoding="utf-8-sig") as f:
            f.write(message)
    except Exception:
        # ログ書き出し自体が失敗しても本来のエラーを優先する
        pass


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit as e:
        # main() の正常終了(0)はそのまま。異常系メッセージはログにも残す
        if e.code not in (0, None):
            detail = str(e.code)
            print(f"ERROR: {detail}", file=sys.stderr)
            _write_error_log(detail)
        raise
    except Exception:
        detail = traceback.format_exc()
        print(detail, file=sys.stderr)
        _write_error_log(detail)
        raise SystemExit(1)
