"""人事成本系統.xlsx 的欄位定義、解析與報表組裝。"""

from __future__ import annotations

import re
from io import BytesIO
from typing import List

import pandas as pd

# --- 選項（依範本說明） ---
YEAR_OPTIONS = [str(y) for y in range(113, 121)]
COMPANY_OPTIONS = ["得意佳", "匯鴻", "寶得", "得威", "鴻源", "馥鼎", "賦潤"]
PROJECT_OPTIONS = [
    "麗寶鐸藝",
    "天水一墅",
    "新潤世界都心",
    "名軒心城市",
    "麗寶之丘",
    "宏樸如嶼",
    "首學杭州",
    "首御臨沂",
    "商用不動產",
    "總公司",
]
HEADQUARTERS_PROJECT = "總公司"
SITE_PROJECT_OPTIONS = [p for p in PROJECT_OPTIONS if p != HEADQUARTERS_PROJECT]
HQ_CASE_FIELD_OPTIONS = ["營收", "營收(未進帳)"]
CASE_FIELD_OPTIONS = ["總銷", "簽約金額", "銷售請款額", "請款淨額", "營收", "營收(未進帳)"]
HR_ITEM_OPTIONS = ["勞保", "勞退", "健保", "二代", "所得稅", "執行業務所得", "薪資", "三節", "獎金", "員工福利"]
HR_MANUAL_ITEM_OPTIONS = ["薪資", "三節", "獎金", "員工福利"]
BONUS_TYPE_OPTIONS = ["個獎", "控獎", "裝個", "裝控", "端午", "中秋", "年終", "績效", "租賃", "特獎"]

CASE_TOTAL_COLS = [
    "年度",
    "公司名",
    "案場",
    "總銷",
    "簽約金額",
    "銷售請款額",
    "請款額1%",
    "請款淨額",
    "營收",
    "營收(未進帳)",
    "人事成本",
    "比例",
]
HR_COST_COLS = ["年度", "案場", "勞保", "勞退", "健保", "二代", "薪資", "三節", "獎金", "員工福利", "總計"]
YEARLY_STAT_COLS = ["姓名", "113年", "114年", "115年", "總計"]
PERSONAL_INCOME_COLS = ["年度", "案場", "姓名", "金額", "所得稅", "執行業務所得", "二代健保", "實領金額"]
MONTHLY_TOTAL_COLS = ["年度", "案場", "項目", "一月", "二月", "三月", "四月", "五月", "六月", "七月", "八月", "九月", "十月", "十一月", "十二月", "總計"]
MONTH_LABELS = MONTHLY_TOTAL_COLS[3:15]
MONTH_NUM_TO_LABEL = {i + 1: label for i, label in enumerate(MONTH_LABELS)}
MONTHLY_TOTAL_ITEMS = ["薪資", "獎金"]

CASE_NOTE_KEYS = ["總銷", "簽約金額", "銷售請款額", "請款額1%", "請款淨額", "營收", "營收(未進帳)"]
CASE_OVERWRITE_FIELDS = {"總銷"}
CASE_DELTA_FIELDS = {"營收(未進帳)"}
HR_NOTE_KEYS = ["勞保", "勞退", "健保", "二代", "薪資", "三節", "獎金", "員工福利", "所得稅", "執行業務所得"]

_STRUCTURED_NOTE_KEYS = frozenset(
    {
        "date",
        "field",
        "category",
        "migrated_from",
        "legacy_id",
        "獎項",
        "次數",
        "請款額1%",
        *HR_NOTE_KEYS,
    }
)
HR_COST_SUM_KEYS = ["勞保", "勞退", "健保", "二代", "薪資", "三節", "獎金", "員工福利"]
HR_IMPORT_DISPLAY_COLS = [
    "年度",
    "案名",
    "姓名",
    "日期",
    "項目",
    "金額",
    "勞保",
    "勞退",
    "保費",
    "金額",
    "稅款",
    "金額",
    "獎項",
    "次數",
    "備註",
]
HR_DETAIL_SHEET_NAMES = ["檔案匯入資料", "人事成本匯入", "人事成本明細"]

LEGACY_CASE_SOURCES = {"薪資占比", "總表主表手動", "總表主表調整"}
LEGACY_HR_SOURCES = {"個人所得", "個人所得手動", "薪資獎金統計", "在職年手動", "在職年調整"}
NEW_CASE_SOURCES = {"全案總表", "全案總表手動", "全案總表調整"}
NEW_HR_SOURCES = {"人事成本", "人事成本手動", "人事成本調整"}
REPORT_VISIBLE_SOURCES = NEW_CASE_SOURCES | NEW_HR_SOURCES


def is_report_visible_source(source_type: object) -> bool:
    return str(source_type or "") in REPORT_VISIBLE_SOURCES


def to_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip().replace(",", "").replace("%", "")
    if text in {"", "-", "nan", "None", "(總計)"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def append_note_parts(parts: list[str]) -> str:
    return ";".join([p for p in parts if str(p).strip()])


def parse_note_value(note: object, key: str) -> str:
    text = "" if pd.isna(note) else str(note)
    match = re.search(rf"{re.escape(key)}:([^;]+)", text)
    return match.group(1).strip() if match else ""


def parse_note_number(note: object, key: str) -> float:
    return to_number(parse_note_value(note, key))


def parse_note_remark(note: object) -> str:
    """取出 note 中非結構化欄位的備註文字。"""
    text = "" if pd.isna(note) else str(note)
    if not text.strip():
        return ""
    parts: list[str] = []
    for segment in text.split(";"):
        seg = segment.strip()
        if not seg:
            continue
        if ":" in seg:
            key = seg.split(":", 1)[0].strip()
            if key in _STRUCTURED_NOTE_KEYS:
                continue
        parts.append(seg)
    return ";".join(parts)


def rebuild_note_display_fields(
    note: object,
    award: object,
    times: object,
    remark: object,
) -> str:
    text = "" if pd.isna(note) else str(note)
    kept: list[str] = []
    for segment in text.split(";"):
        seg = segment.strip()
        if not seg:
            continue
        if seg.startswith("獎項:") or seg.startswith("次數:"):
            continue
        if ":" in seg:
            key = seg.split(":", 1)[0].strip()
            if key in _STRUCTURED_NOTE_KEYS:
                kept.append(seg)
    award_text = clean_text(award)
    times_text = clean_text(times) or parse_note_value(text, "次數") or "1"
    remark_text = clean_text(remark)
    if award_text:
        kept.append(f"獎項:{award_text}")
    kept.append(f"次數:{times_text}")
    if remark_text:
        kept.append(remark_text)
    return ";".join(kept)


def roc_year_from_value(value: object) -> int | None:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text or text in {"年度", "nan"}:
        return None
    num = to_number(text)
    if 100 <= num <= 200:
        return int(num)
    return None


def calc_request_pct(sales_request: float) -> float:
    return round(float(sales_request or 0) * 0.01)


def calc_hr_ratio(hr_cost: float, request_pct: float) -> float:
    if request_pct <= 0:
        return 0.0
    return round((hr_cost / request_pct) * 100, 4)


def is_headquarters_project(project: object) -> bool:
    return clean_text(project) == HEADQUARTERS_PROJECT


def hq_revenue_base(revenue: float, revenue_pending: float) -> float:
    return float(revenue or 0) + float(revenue_pending or 0)


def calc_hq_hr_ratio(hr_cost: float, revenue: float, revenue_pending: float) -> float:
    base = hq_revenue_base(revenue, revenue_pending)
    if base <= 0:
        return 0.0
    return round((float(hr_cost or 0) / base) * 100, 4)


def case_row_hr_ratio(row: pd.Series | dict) -> float:
    hr_cost = float(row.get("人事成本") or 0)
    if is_headquarters_project(row.get("案場")):
        return calc_hq_hr_ratio(
            hr_cost,
            float(row.get("營收") or 0),
            float(row.get("營收(未進帳)") or 0),
        )
    request_pct = float(row.get("請款額1%") or 0)
    if request_pct > 0:
        return calc_hr_ratio(hr_cost, request_pct)
    return float(row.get("比例") or 0)


def parse_year_amounts(note: object) -> dict[str, float]:
    text = "" if pd.isna(note) else str(note)
    pairs = re.findall(r"(\d{2,4})\s*:\s*([-\d.]+)", text)
    out: dict[str, float] = {}
    for y, v in pairs:
        out[y] = to_number(v)
    return out


def parse_manual_category(note: object) -> str:
    text = "" if pd.isna(note) else str(note)
    match = re.search(r"category:([^;]+)", text)
    return match.group(1).strip() if match else ""


def is_migrated_note(note: object) -> bool:
    return "migrated_from:" in str(note or "")


def is_legacy_hr_source(source_type: object) -> bool:
    st = str(source_type or "")
    if st in LEGACY_HR_SOURCES:
        return True
    return st.startswith("個人所得調整_") or "總表分表" in st


def is_legacy_case_source(source_type: object) -> bool:
    return str(source_type or "") in LEGACY_CASE_SOURCES


def list_legacy_source_counts(df_all: pd.DataFrame) -> dict[str, int]:
    if df_all.empty:
        return {}
    counts: dict[str, int] = {}
    for _, row in df_all.iterrows():
        if is_migrated_note(row.get("note")):
            continue
        st = str(row.get("source_type") or "")
        if st in NEW_CASE_SOURCES or st in NEW_HR_SOURCES:
            continue
        if is_legacy_case_source(st) or is_legacy_hr_source(st):
            counts[st] = counts.get(st, 0) + 1
    return counts


def delete_legacy_data() -> tuple[int, dict[str, int]]:
    """刪除舊版來源資料（薪資占比、個人所得、總表分表等），保留新格式資料。"""
    from database import delete_records_by_source, delete_records_by_source_like

    stats: dict[str, int] = {}
    total = 0
    for source_type in sorted(LEGACY_CASE_SOURCES | LEGACY_HR_SOURCES):
        deleted = delete_records_by_source(source_type)
        if deleted:
            stats[source_type] = deleted
            total += deleted
    for pattern in ["個人所得調整_%", "%總表分表%"]:
        deleted = delete_records_by_source_like(pattern)
        if deleted:
            stats[pattern] = deleted
            total += deleted
    return total, stats


def delete_non_report_data() -> tuple[int, dict[str, int]]:
    """刪除不會出現在報表的舊來源（如總表分表匯入、薪資占比、legacy_migration）。"""
    from database import delete_records_by_source, get_connection
    from contextlib import closing

    stats: dict[str, int] = {}
    total = 0
    with closing(get_connection()) as conn:
        types = [str(r[0]) for r in conn.execute("SELECT DISTINCT source_type FROM payroll_records").fetchall()]
    for source_type in types:
        if is_report_visible_source(source_type):
            continue
        deleted = delete_records_by_source(source_type)
        if deleted:
            stats[source_type] = deleted
            total += deleted
    return total, stats


def clear_hr_report_data() -> tuple[int, dict[str, int]]:
    """匯入前清除人事成本相關資料（新格式 + 舊版）。"""
    from database import delete_records_by_source, delete_records_by_source_like

    stats: dict[str, int] = {}
    total = 0
    for source_type in sorted(NEW_HR_SOURCES | LEGACY_HR_SOURCES):
        deleted = delete_records_by_source(source_type)
        if deleted:
            stats[source_type] = deleted
            total += deleted
    for pattern in ["個人所得調整_%", "%總表分表%"]:
        deleted = delete_records_by_source_like(pattern)
        if deleted:
            stats[pattern] = deleted
            total += deleted
    return total, stats


def clear_case_report_data() -> tuple[int, dict[str, int]]:
    """匯入前清除全案總表相關資料（新格式 + 舊版）。"""
    from database import delete_records_by_source

    stats: dict[str, int] = {}
    total = 0
    for source_type in sorted(NEW_CASE_SOURCES | LEGACY_CASE_SOURCES):
        deleted = delete_records_by_source(source_type)
        if deleted:
            stats[source_type] = deleted
            total += deleted
    return total, stats


def _normalize_case_note(note: object) -> str:
    text = "" if pd.isna(note) else str(note)
    replacements = {
        "全案總銷": "總銷",
        "記帳": "營收",
    }
    for old, new in replacements.items():
        text = re.sub(rf"{re.escape(old)}:", f"{new}:", text)
    return text


def _hr_display_item(year: int, project: str, note: str, total: float) -> dict:
    return {
        "年度": year,
        "案場": project,
        "勞保": parse_note_number(note, "勞保"),
        "勞退": parse_note_number(note, "勞退"),
        "健保": parse_note_number(note, "健保"),
        "二代": parse_note_number(note, "二代"),
        "薪資": parse_note_number(note, "薪資"),
        "三節": parse_note_number(note, "三節"),
        "獎金": parse_note_number(note, "獎金"),
        "員工福利": parse_note_number(note, "員工福利"),
        "總計": total or sum(parse_note_number(note, k) for k in HR_NOTE_KEYS),
    }


def legacy_hr_display_items(row: dict) -> list[dict]:
    source_type = str(row.get("source_type") or "")
    note = str(row.get("note") or "")
    project = str(row.get("project_name") or "")
    year = roc_year_from_value(row.get("roc_year")) or 114

    if source_type in {"個人所得", "個人所得手動"} or source_type.startswith("個人所得調整"):
        gross = float(row.get("salary") or 0)
        income_tax = parse_note_number(note, "所得稅")
        health2 = parse_note_number(note, "二代")
        business = parse_note_number(note, "執行業務所得")
        if income_tax == 0:
            income_tax = float(row.get("bonus") or 0)
        if health2 == 0:
            health2 = float(row.get("welfare") or 0)
        if business == 0 and parse_note_value(note, "所得類型") == "執行業務所得":
            business = income_tax
            income_tax = 0.0
        hr_note = append_note_parts(
            [
                f"薪資:{gross}",
                "獎金:0",
                f"所得稅:{income_tax}",
                f"二代:{health2}",
                f"執行業務所得:{business}",
            ]
        )
        return [_hr_display_item(year, project, hr_note, gross)]

    if source_type in {"薪資獎金統計", "在職年手動", "在職年調整"}:
        items: list[dict] = []
        year_amounts = parse_year_amounts(note)
        if not year_amounts:
            amt = float(row.get("bonus") or 0) + float(row.get("salary") or 0)
            if amt > 0:
                year_amounts[str(year)] = amt
        for y_text, amt in year_amounts.items():
            y = roc_year_from_value(y_text) or year
            hr_note = append_note_parts([f"薪資:0", f"獎金:{amt}"])
            items.append(_hr_display_item(y, project, hr_note, amt))
        return items

    if "總表分表" in source_type:
        category = parse_manual_category(note)
        if category not in HR_NOTE_KEYS:
            category = category if category in HR_MANUAL_ITEM_OPTIONS else "薪資"
        amount = float(row.get("total_income") or 0)
        hr_note = append_note_parts([f"{category}:{amount}"])
        return [_hr_display_item(year, project, hr_note, amount)]

    return []


def legacy_personal_income_items(row: dict) -> list[dict]:
    source_type = str(row.get("source_type") or "")
    if source_type not in {"個人所得", "個人所得手動"} and not source_type.startswith("個人所得調整"):
        return []
    year = roc_year_from_value(row.get("roc_year")) or 114
    name = clean_text(row.get("employee_name"))
    if not name:
        return []
    note = str(row.get("note") or "")
    amount = float(row.get("salary") or 0)
    income_tax = parse_note_number(note, "所得稅") or float(row.get("bonus") or 0)
    health2 = parse_note_number(note, "二代") or float(row.get("welfare") or 0)
    business = parse_note_number(note, "執行業務所得")
    if business == 0 and parse_note_value(note, "所得類型") == "執行業務所得":
        business = income_tax
        income_tax = 0.0
    net = float(row.get("total_income") or 0) or round(amount - income_tax - business - health2)
    return [
        {
            "年度": year,
            "案場": str(row.get("project_name") or ""),
            "姓名": name,
            "金額": amount,
            "所得稅": income_tax,
            "執行業務所得": business,
            "二代健保": health2,
            "實領金額": net,
        }
    ]


def migrate_legacy_row(row: dict) -> list[tuple[str, dict]]:
    """回傳 [(新來源類型, 紀錄dict), ...]"""
    if is_migrated_note(row.get("note")):
        return []

    source_type = str(row.get("source_type") or "")
    legacy_id = row.get("id")
    note = str(row.get("note") or "")
    base_note = append_note_parts([f"migrated_from:{source_type}", f"legacy_id:{legacy_id}", _normalize_case_note(note)])

    if source_type == "薪資占比":
        mapped_note = _normalize_case_note(note)
        if "總銷:" not in mapped_note and parse_note_number(note, "全案總銷") > 0:
            mapped_note = append_note_parts([mapped_note, f"總銷:{parse_note_number(note, '全案總銷')}"])
        return [
            (
                "全案總表",
                {
                    "sheet_name": "全案總表",
                    "employee_name": None,
                    "company_name": row.get("company_name"),
                    "project_name": row.get("project_name"),
                    "roc_year": roc_year_from_value(row.get("roc_year")) or 114,
                    "salary": float(row.get("salary") or 0),
                    "bonus": float(row.get("bonus") or 0),
                    "welfare": float(row.get("welfare") or 0),
                    "total_income": float(row.get("total_income") or 0),
                    "note": append_note_parts([base_note, mapped_note]),
                },
            )
        ]

    if source_type in {"總表主表手動", "總表主表調整"}:
        target = "全案總表手動" if source_type == "總表主表手動" else "全案總表調整"
        labor = parse_note_number(note, "勞健退")
        festival = parse_note_number(note, "三節")
        meal = parse_note_number(note, "餐費")
        mapped = append_note_parts(
            [
                base_note,
                f"勞保:{labor}",
                f"三節:{festival}",
                f"餐費:{meal}",
                f"薪資:{float(row.get('salary') or 0)}",
                f"獎金:{float(row.get('bonus') or 0)}",
                f"員工福利:{float(row.get('welfare') or 0)}",
            ]
        )
        return [
            (
                target,
                {
                    "sheet_name": "全案總表",
                    "employee_name": "未指定",
                    "company_name": row.get("company_name"),
                    "project_name": row.get("project_name"),
                    "roc_year": roc_year_from_value(row.get("roc_year")) or 114,
                    "salary": float(row.get("salary") or 0),
                    "bonus": float(row.get("bonus") or 0),
                    "welfare": float(row.get("welfare") or 0),
                    "total_income": float(row.get("total_income") or 0),
                    "note": mapped,
                },
            )
        ]

    if is_legacy_hr_source(source_type):
        items = legacy_hr_display_items(row)
        records: list[tuple[str, dict]] = []
        for item in items:
            hr_note = append_note_parts(
                [
                    base_note,
                    *[f"{k}:{item[k]}" for k in HR_NOTE_KEYS if item.get(k)],
                    f"所得稅:{parse_note_number(note, '所得稅')}",
                    f"執行業務所得:{parse_note_number(note, '執行業務所得')}",
                ]
            )
            records.append(
                (
                    "人事成本",
                    {
                        "sheet_name": "人事成本",
                        "employee_name": clean_text(row.get("employee_name")) or None,
                        "company_name": row.get("company_name"),
                        "project_name": row.get("project_name") or item.get("案場"),
                        "roc_year": item.get("年度"),
                        "salary": item.get("薪資", 0),
                        "bonus": item.get("獎金", 0),
                        "welfare": item.get("員工福利", 0),
                        "total_income": item.get("總計", 0),
                        "note": hr_note,
                    },
                )
            )
        return records

    return []


def migrate_all_legacy_records(rows: list[dict]) -> tuple[list[tuple[str, dict]], dict[str, int]]:
    migrated_legacy_ids = {
        parse_note_value(r.get("note"), "legacy_id")
        for r in rows
        if is_migrated_note(r.get("note"))
    }
    migrated_legacy_ids.discard("")

    output: list[tuple[str, dict]] = []
    stats: dict[str, int] = {}
    for row in rows:
        if is_migrated_note(row.get("note")):
            continue
        legacy_id = str(row.get("id") or "")
        if legacy_id and legacy_id in migrated_legacy_ids:
            continue
        converted = migrate_legacy_row(row)
        if not converted:
            continue
        output.extend(converted)
        st = str(row.get("source_type") or "")
        stats[st] = stats.get(st, 0) + 1
    return output, stats


def hr_row_total(note: object) -> float:
    return sum(parse_note_number(note, k) for k in HR_NOTE_KEYS)


def hr_cost_sum_from_note(note: object) -> float:
    return sum(parse_note_number(note, k) for k in HR_COST_SUM_KEYS)


def normalize_insurance_type(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if "二代" in text:
        return "二代"
    if "健" in text:
        return "健保"
    return text


def normalize_tax_type(value: object) -> str:
    text = clean_text(value).replace(" ", "")
    if not text:
        return ""
    # 執行業務須先於「所得」，避免「執行業務所得」被誤判
    if any(k in text for k in ("執行業務所得", "執行業務", "執業務", "執業")):
        return "執行業務所得"
    if "執行" in text and "業" in text:
        return "執行業務所得"
    if "業務" in text:
        return "執行業務所得"
    if "所得" in text:
        return "所得稅"
    return ""


def parse_tax_amounts_from_note(note: object) -> tuple[float, float, float]:
    """從 note 讀取 (所得稅, 執行業務所得, 二代健保)。"""
    return (
        parse_note_number(note, "所得稅"),
        parse_note_number(note, "執行業務所得"),
        parse_note_number(note, "二代") or parse_note_number(note, "二代健保"),
    )


def parse_import_row_taxes(row: pd.Series) -> tuple[float, float, float]:
    """從匯入列讀取稅務金額。"""
    income_tax = to_number(row.get("所得稅")) if "所得稅" in row.index else 0.0
    business = to_number(row.get("執行業務所得")) if "執行業務所得" in row.index else 0.0
    health2 = to_number(row.get("二代健保")) if "二代健保" in row.index else 0.0

    tax_raw = import_row_tax_label(row)
    tax_amount = to_number(row.get("稅款金額"))
    if tax_amount > 0:
        kind = normalize_tax_type(tax_raw)
        if kind == "所得稅":
            income_tax = tax_amount
        elif kind == "執行業務所得":
            business = tax_amount
        elif tax_raw:
            # 有稅別文字但無法辨識時，不預設塞進所得稅
            pass

    ins_type = normalize_insurance_type(row.get("保費"))
    ins_amount = to_number(row.get("保費金額"))
    if ins_type == "二代" and ins_amount > 0:
        health2 = ins_amount
    return income_tax, business, health2


def parse_roc_date_text(value: object) -> str:
    text = clean_text(value)
    if not text:
        return ""
    if re.match(r"\d{4}-\d{2}-\d{2}", text):
        return text
    m = re.match(r"(\d{2,3})[/\-](\d{1,2})[/\-](\d{1,2})", text)
    if m:
        roc = int(m.group(1))
        if roc < 1911:
            roc += 1911
        return f"{roc:04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    try:
        ts = pd.to_datetime(value, errors="coerce")
        if pd.notna(ts):
            return ts.date().isoformat()
    except Exception:
        pass
    return text


def payroll_report_month_from_date(value: object) -> tuple[int, int] | None:
    """薪資月份：M 月資料對應日期為 (M+1) 月 1 日～25 日（例：1 月 = 2/1～2/25）。"""
    text = parse_roc_date_text(value)
    if not text:
        return None
    m = re.match(r"(\d+)-(\d+)-(\d+)", text)
    if not m:
        return None
    year_raw = int(m.group(1))
    cal_month = int(m.group(2))
    day = int(m.group(3))
    if cal_month < 1 or cal_month > 12 or day < 1 or day > 25:
        return None
    roc_year = year_raw - 1911 if year_raw >= 1911 else year_raw
    if cal_month == 1:
        return roc_year - 1, 12
    return roc_year, cal_month - 1


def _clean_import_col(name: object) -> str:
    return clean_text(name).replace("\n", "")


TAX_LABEL_ALIASES = {"稅務", "稅別", "稅種", "稅目"}


def _canonical_import_col(col: str) -> str:
    if col in TAX_LABEL_ALIASES:
        return "稅款"
    return col


def import_row_tax_label(row: pd.Series) -> str:
    for key in ("稅款", "稅務", "稅別", "稅種", "稅目"):
        if key in row.index:
            text = clean_text(row.get(key))
            if text:
                return text
    return ""


def _rename_hr_import_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    raw = [_canonical_import_col(_clean_import_col(c)) for c in out.columns]
    new_cols: list[str] = []
    amount_idx = 0
    for i, col in enumerate(raw):
        if not col:
            new_cols.append(f"欄位{i + 1}")
            continue
        if col == "金額" or re.match(r"^金額\.\d+$", col):
            labels = ["項目金額", "保費金額", "稅款金額"]
            new_cols.append(labels[amount_idx] if amount_idx < len(labels) else f"金額{amount_idx + 1}")
            amount_idx += 1
        else:
            new_cols.append(col)
    out.columns = new_cols
    if "案場" in out.columns and "案名" not in out.columns:
        out = out.rename(columns={"案場": "案名"})
    return out


def _forward_fill_import_group_keys(df: pd.DataFrame) -> pd.DataFrame:
    """Excel 合併儲存格時，姓名/日期等常只有第一列有值，向下沿用。"""
    out = df.copy()
    for col in ("年度", "案名", "姓名", "日期"):
        if col not in out.columns:
            continue
        series = out[col].map(lambda x: clean_text(x) if not pd.isna(x) else "")
        series = series.mask(series == "", pd.NA).ffill().fillna("")
        out[col] = series
    return out


def is_hr_detail_import_df(df: pd.DataFrame) -> bool:
    if df.empty:
        return False
    renamed = _rename_hr_import_columns(df)
    cols = set(renamed.columns)
    return "姓名" in cols and "項目" in cols and ("案名" in cols or "案場" in cols)


def build_hr_cost_record(
    *,
    roc_year: int,
    project_name: str,
    employee_name: str,
    date_text: str,
    item: str,
    item_amount: float,
    labor: float,
    pension: float,
    insurance_type: str,
    insurance_amount: float,
    income_tax: float = 0.0,
    business_income: float = 0.0,
    health2: float = 0.0,
    bonus_type: str = "",
    times: str = "1",
    remark: str = "",
) -> dict:
    item_amounts = {name: 0.0 for name in HR_MANUAL_ITEM_OPTIONS}
    if item in item_amounts:
        item_amounts[item] = float(item_amount or 0)
    elif float(item_amount or 0) > 0:
        # 項目空白或非薪資/獎金/三節/員工福利時，金額仍計入薪資
        item_amounts["薪資"] = float(item_amount or 0)

    ins_type = normalize_insurance_type(insurance_type)
    health = float(insurance_amount or 0) if ins_type == "健保" else 0.0
    nhi2 = float(health2 or 0) if health2 else (float(insurance_amount or 0) if ins_type == "二代" else 0.0)
    income_tax = float(income_tax or 0)
    business = float(business_income or 0)

    note_parts = [
        f"date:{date_text}" if date_text else "",
        f"勞保:{float(labor or 0)}",
        f"勞退:{float(pension or 0)}",
        f"健保:{health}",
        f"二代:{nhi2}",
        f"所得稅:{income_tax}",
        f"執行業務所得:{business}",
        f"薪資:{item_amounts['薪資']}",
        f"三節:{item_amounts['三節']}",
        f"獎金:{item_amounts['獎金']}",
        f"員工福利:{item_amounts['員工福利']}",
        f"次數:{times or '1'}",
    ]
    if bonus_type.strip():
        note_parts.append(f"獎項:{bonus_type.strip()}")
    if remark.strip():
        note_parts.append(remark.strip())
    note = append_note_parts(note_parts)
    total = hr_cost_sum_from_note(note)
    return {
        "sheet_name": "人事成本",
        "employee_name": employee_name,
        "company_name": None,
        "project_name": project_name,
        "roc_year": roc_year,
        "salary": item_amounts["薪資"],
        "bonus": item_amounts["獎金"],
        "welfare": item_amounts["員工福利"],
        "total_income": total,
        "note": note,
    }


def _hr_detail_row_to_record(row: pd.Series) -> dict | None:
    year = roc_year_from_value(row.get("年度"))
    project = clean_text(row.get("案名"))
    name = clean_text(row.get("姓名"))
    if year is None or not project:
        return None
    if _is_instruction_row(row.get("年度"), project, name, row.get("項目")):
        return None

    item = clean_text(row.get("項目"))
    item_amount = to_number(row.get("項目金額"))
    if not item and item_amount > 0:
        item = "薪資"
    labor = to_number(row.get("勞保"))
    pension = to_number(row.get("勞退"))
    insurance_type = normalize_insurance_type(row.get("保費"))
    insurance_amount = to_number(row.get("保費金額"))
    income_tax, business_income, health2 = parse_import_row_taxes(row)
    bonus_type = clean_text(row.get("獎項"))
    times = clean_text(row.get("次數")) or "1"
    remark = clean_text(row.get("備註"))
    date_text = parse_roc_date_text(row.get("日期"))

    if not name and not any(
        [
            item_amount,
            labor,
            pension,
            insurance_amount,
            income_tax,
            business_income,
            health2,
            item,
        ]
    ):
        return None

    return build_hr_cost_record(
        roc_year=year,
        project_name=project,
        employee_name=name or "未命名",
        date_text=date_text,
        item=item,
        item_amount=item_amount,
        labor=labor,
        pension=pension,
        insurance_type=insurance_type,
        insurance_amount=insurance_amount,
        income_tax=income_tax,
        business_income=business_income,
        health2=health2,
        bonus_type=bonus_type,
        times=times,
        remark=remark,
    )


def parse_hr_detail_dataframe(df: pd.DataFrame) -> tuple[list[dict], pd.DataFrame]:
    renamed = _forward_fill_import_group_keys(_rename_hr_import_columns(df))
    records: list[dict] = []
    preview_rows: list[dict] = []
    for _, row in renamed.iterrows():
        record = _hr_detail_row_to_record(row)
        if record is None:
            continue
        records.append(record)
        preview_rows.append(
            {
                "年度": record["roc_year"],
                "案名": record["project_name"],
                "姓名": record["employee_name"],
                "日期": parse_note_value(record["note"], "date") or parse_roc_date_text(row.get("日期")),
                "項目": clean_text(row.get("項目")),
                "項目金額": to_number(row.get("項目金額")),
                "勞保": parse_note_number(record["note"], "勞保"),
                "勞退": parse_note_number(record["note"], "勞退"),
                "保費": clean_text(row.get("保費")),
                "保費金額": to_number(row.get("保費金額")),
                "稅款": import_row_tax_label(row),
                "稅款金額": to_number(row.get("稅款金額")),
                "所得稅": parse_note_number(record["note"], "所得稅"),
                "執行業務所得": parse_note_number(record["note"], "執行業務所得"),
                "獎項": parse_note_value(record["note"], "獎項"),
                "次數": parse_note_value(record["note"], "次數"),
                "備註": clean_text(row.get("備註")),
            }
        )
    preview = pd.DataFrame(preview_rows)
    return records, preview


def _read_import_sheets(file_bytes: bytes, filename: str = "") -> list[tuple[str, pd.DataFrame]]:
    name = (filename or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(BytesIO(file_bytes), encoding="utf-8-sig")
        return [("csv", df)]
    xls = pd.ExcelFile(BytesIO(file_bytes))
    sheets: list[tuple[str, pd.DataFrame]] = []
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=0)
        sheets.append((sheet_name, df))
    return sheets


def parse_hr_detail_workbook(file_bytes: bytes, filename: str = "") -> tuple[list[dict], pd.DataFrame]:
    sheets = _read_import_sheets(file_bytes, filename)
    candidates: list[tuple[int, int, str, pd.DataFrame]] = []
    for idx, (sheet_name, df) in enumerate(sheets):
        if sheet_name in HR_DETAIL_SHEET_NAMES or is_hr_detail_import_df(df):
            priority = 0 if sheet_name in HR_DETAIL_SHEET_NAMES else 1
            candidates.append((priority, idx, sheet_name, df))
    candidates.sort()
    for _, _, _sheet_name, df in candidates:
        records, preview = parse_hr_detail_dataframe(df)
        if records:
            return records, preview
    return [], pd.DataFrame()


def build_hr_import_template_bytes() -> bytes:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "檔案匯入資料"
    ws.append(HR_IMPORT_DISPLAY_COLS)
    ws.append(
        [
            "114",
            "天水一墅",
            "範例姓名",
            "115/05/10",
            "薪資",
            50000,
            1000,
            200,
            "健保",
            500,
            "所得稅",
            0,
            "個獎",
            "1",
            "",
        ]
    )
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _is_instruction_row(*cells: object) -> bool:
    joined = " ".join(clean_text(c) for c in cells)
    return any(x in joined for x in ["資料條件", "手動新增", "需求", "項目", "(總計)"])


def _row_amount(row: pd.Series, *column_names: str) -> float:
    for name in column_names:
        if name in row.index:
            return to_number(row.get(name))
    return 0.0


def _case_note_from_row(row: pd.Series) -> str:
    sales_request = _row_amount(row, "銷售請款額")
    request_pct = _row_amount(row, "請款額1%")
    if request_pct == 0 and sales_request > 0:
        request_pct = calc_request_pct(sales_request)
    return append_note_parts(
        [
            f"總銷:{_row_amount(row, '總銷')}",
            f"簽約金額:{_row_amount(row, '簽約金額')}",
            f"銷售請款額:{sales_request}",
            f"請款額1%:{request_pct}",
            f"請款淨額:{_row_amount(row, '請款淨額')}",
            f"營收:{_row_amount(row, '營收')}",
            f"營收(未進帳):{_row_amount(row, '營收(未進帳)', '營收\\n(未進帳)')}",
        ]
    )


def _hr_note_from_row(row: pd.Series) -> str:
    parts = [f"{k}:{to_number(row.get(k))}" for k in HR_NOTE_KEYS if k in row.index]
    return append_note_parts(parts)


def parse_hr_system_workbook(file_bytes: bytes) -> dict[str, List[dict]]:
    bio = BytesIO(file_bytes)
    xls = pd.ExcelFile(bio)
    result: dict[str, List[dict]] = {"全案總表": [], "人事成本": []}

    if "全案總表" in xls.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name="全案總表", header=0)
        df.columns = [clean_text(c).replace("\n", "") for c in df.columns]
        for _, row in df.iterrows():
            year = roc_year_from_value(row.get("年度"))
            company = clean_text(row.get("公司名"))
            project = clean_text(row.get("案場"))
            if year is None or not company or not project:
                continue
            if _is_instruction_row(row.get("年度"), company, project):
                continue
            note = _case_note_from_row(row)
            hr_cost = _row_amount(row, "人事成本")
            ratio = _row_amount(row, "比例", "比例(人事/請款額1%)*100%")
            result["全案總表"].append(
                {
                    "sheet_name": "全案總表",
                    "employee_name": None,
                    "company_name": company,
                    "project_name": project,
                    "roc_year": year,
                    "salary": 0.0,
                    "bonus": 0.0,
                    "welfare": 0.0,
                    "total_income": hr_cost,
                    "note": append_note_parts([note, f"人事成本:{hr_cost}", f"比例:{ratio}"]),
                }
            )

    for sheet_name in HR_DETAIL_SHEET_NAMES:
        if sheet_name in xls.sheet_names:
            df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=0)
            detail_records, _ = parse_hr_detail_dataframe(df)
            result["人事成本"].extend(detail_records)

    if "人事成本" in xls.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name="人事成本", header=0)
        if is_hr_detail_import_df(df):
            detail_records, _ = parse_hr_detail_dataframe(df)
            result["人事成本"].extend(detail_records)
        else:
            df.columns = [clean_text(c).replace("\n", "") for c in df.columns]
            for _, row in df.iterrows():
                year = roc_year_from_value(row.get("年度"))
                project = clean_text(row.get("案場"))
                if year is None or not project:
                    continue
                if _is_instruction_row(row.get("年度"), project):
                    continue
                note = _hr_note_from_row(row)
                total = to_number(row.get("總計"))
                if total <= 0:
                    total = hr_cost_sum_from_note(note)
                result["人事成本"].append(
                    {
                        "sheet_name": "人事成本",
                        "employee_name": clean_text(row.get("姓名")) or None,
                        "company_name": None,
                        "project_name": project,
                        "roc_year": year,
                        "salary": parse_note_number(note, "薪資"),
                        "bonus": parse_note_number(note, "獎金"),
                        "welfare": parse_note_number(note, "員工福利"),
                        "total_income": total,
                        "note": note,
                    }
                )

    return result


def _filter_sources(df_all: pd.DataFrame, names: list[str]) -> pd.DataFrame:
    if df_all.empty:
        return df_all
    return df_all[df_all["source_type"].isin(names)].copy()


def _merge_case_notes(note_series: pd.Series) -> dict[str, float]:
    summed = {key: 0.0 for key in CASE_NOTE_KEYS}
    overwrite: dict[str, float] = {}
    extras = {"人事成本": 0.0, "比例": 0.0}

    for note in note_series:
        text = "" if pd.isna(note) else str(note)
        field = parse_note_value(text, "field")
        mode = parse_note_value(text, "mode")

        for key in CASE_NOTE_KEYS:
            val = parse_note_number(text, key)
            if val == 0 and field and field != key:
                continue
            if key in CASE_OVERWRITE_FIELDS or (mode == "overwrite" and (not field or field == key)):
                overwrite[key] = val
            else:
                summed[key] += val

        for key in ["人事成本", "比例"]:
            extras[key] += parse_note_number(text, key)

    merged = {**summed, **overwrite}
    for key, val in extras.items():
        if val != 0:
            merged[key] = merged.get(key, 0.0) + val
    return merged


def build_case_total_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    sources = _filter_sources(df_all, list(NEW_CASE_SOURCES))

    hr_agg = build_hr_cost_frame(df_all, filter_year=filter_year)
    hr_lookup: dict[tuple[int, str], float] = {}
    if not hr_agg.empty:
        for _, r in hr_agg.iterrows():
            yr = roc_year_from_value(r.get("年度"))
            if yr is None:
                continue
            hr_lookup[(yr, str(r.get("案場") or ""))] = float(r.get("總計") or 0)

    if sources.empty and not hr_lookup:
        return pd.DataFrame(columns=CASE_TOTAL_COLS)

    rows: list[dict] = []
    covered_projects: set[tuple[int, str]] = set()

    if not sources.empty:
        grouped = sources.groupby(["roc_year", "company_name", "project_name"], dropna=False)
    else:
        grouped = []

    for (year, company, project), grp in grouped:
        yr = int(year) if pd.notna(year) and to_number(year) >= 100 else None
        if yr is None:
            continue
        if filter_year is not None and yr != filter_year:
            continue
        grp_sorted = grp.sort_values("id") if "id" in grp.columns else grp
        merged = _merge_case_notes(grp_sorted["note"])
        if merged.get("總銷", 0) == 0:
            merged["總銷"] = parse_note_number(";".join(grp_sorted["note"].astype(str)), "全案總銷")
        sales_request = merged.get("銷售請款額", 0.0)
        request_pct = merged.get("請款額1%", 0.0)
        if request_pct == 0 and sales_request > 0:
            request_pct = calc_request_pct(sales_request)
        hr_cost = merged.get("人事成本", 0.0)
        if hr_cost == 0:
            hr_cost = hr_lookup.get((yr, str(project or "")), 0.0)
        ratio = merged.get("比例", 0.0)
        if ratio == 0:
            if is_headquarters_project(project):
                ratio = calc_hq_hr_ratio(
                    hr_cost,
                    merged.get("營收", 0.0),
                    merged.get("營收(未進帳)", 0.0),
                )
            else:
                ratio = calc_hr_ratio(hr_cost, request_pct)
        project_key = str(project or "")
        covered_projects.add((yr, project_key))
        rows.append(
            {
                "年度": yr,
                "公司名": str(company or ""),
                "案場": project_key,
                "總銷": merged.get("總銷", 0.0),
                "簽約金額": merged.get("簽約金額", 0.0),
                "銷售請款額": sales_request,
                "請款額1%": request_pct,
                "請款淨額": merged.get("請款淨額", 0.0),
                "營收": merged.get("營收", 0.0),
                "營收(未進帳)": merged.get("營收(未進帳)", 0.0),
                "人事成本": hr_cost,
                "比例": ratio,
            }
        )

    for (yr, project), hr_cost in hr_lookup.items():
        if filter_year is not None and yr != filter_year:
            continue
        if (yr, project) in covered_projects:
            continue
        rows.append(
            {
                "年度": yr,
                "公司名": "",
                "案場": project,
                "總銷": 0.0,
                "簽約金額": 0.0,
                "銷售請款額": 0.0,
                "請款額1%": 0.0,
                "請款淨額": 0.0,
                "營收": 0.0,
                "營收(未進帳)": 0.0,
                "人事成本": hr_cost,
                "比例": 0.0,
            }
        )

    if not rows:
        return pd.DataFrame(columns=CASE_TOTAL_COLS)
    out = pd.DataFrame(rows)
    grouped = out.groupby(["年度", "公司名", "案場"], as_index=False).sum(numeric_only=True)
    for col in CASE_OVERWRITE_FIELDS:
        if col not in grouped.columns:
            continue
        last_vals = out.groupby(["年度", "公司名", "案場"], as_index=False)[col].last()
        grouped[col] = last_vals[col].values
    grouped["比例"] = grouped.apply(case_row_hr_ratio, axis=1)
    return grouped[CASE_TOTAL_COLS]


def build_hr_cost_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    rows: list[dict] = []
    if df_all.empty:
        return pd.DataFrame(columns=HR_COST_COLS)

    new_sources = _filter_sources(df_all, list(NEW_HR_SOURCES))
    for _, row in new_sources.iterrows():
        yr = roc_year_from_value(row.get("roc_year"))
        if yr is None:
            continue
        if filter_year is not None and yr != filter_year:
            continue
        note = row.get("note")
        item = _hr_display_item(yr, str(row.get("project_name") or ""), str(note or ""), float(row.get("total_income") or 0))
        rows.append(item)

    if not rows:
        return pd.DataFrame(columns=HR_COST_COLS)
    out = pd.DataFrame(rows).groupby(["年度", "案場"], as_index=False).sum(numeric_only=True)
    return out[HR_COST_COLS]


def build_yearly_stat_frame(df_all: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict] = []
    if df_all.empty:
        return pd.DataFrame(columns=YEARLY_STAT_COLS)

    def append_yearly(name: str, year: int, amount: float) -> None:
        if amount <= 0:
            return
        rows.append({"姓名": name, "年度": year, "金額": amount})

    new_sources = _filter_sources(df_all, list(NEW_HR_SOURCES))
    for _, row in new_sources.iterrows():
        name = clean_text(row.get("employee_name"))
        if not name:
            name = f"{row.get('project_name') or '未命名'}-彙總"
        yr = roc_year_from_value(row.get("roc_year"))
        if yr is None:
            continue
        note = row.get("note")
        amount = sum(parse_note_number(note, key) for key in HR_MANUAL_ITEM_OPTIONS)
        if amount <= 0:
            amount = float(row.get("salary") or 0) + float(row.get("bonus") or 0)
        append_yearly(name, yr, amount)

    if not rows:
        return pd.DataFrame(columns=YEARLY_STAT_COLS)

    df = pd.DataFrame(rows)
    pivot = df.pivot_table(index="姓名", columns="年度", values="金額", aggfunc="sum", fill_value=0.0)
    out = pivot.reset_index()
    rename_map = {}
    for col in out.columns:
        if col == "姓名":
            continue
        try:
            rename_map[col] = f"{int(float(col))}年"
        except (TypeError, ValueError):
            rename_map[col] = str(col)
    out = out.rename(columns=rename_map)
    for col in ["113年", "114年", "115年"]:
        if col not in out.columns:
            out[col] = 0.0
    year_cols = [c for c in ["113年", "114年", "115年"] if c in out.columns]
    for col in year_cols:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    out["總計"] = out[year_cols].sum(axis=1) if year_cols else 0.0
    return out[["姓名", *year_cols, "總計"]]


def _personal_income_gross(note: object, row: pd.Series | None = None) -> float:
    gross = sum(parse_note_number(note, key) for key in HR_MANUAL_ITEM_OPTIONS)
    if gross <= 0 and row is not None:
        gross = (
            float(row.get("salary") or 0)
            + float(row.get("bonus") or 0)
            + float(row.get("welfare") or 0)
            + parse_note_number(note, "三節")
        )
    return gross


def build_monthly_total_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    if df_all.empty:
        return pd.DataFrame(columns=MONTHLY_TOTAL_COLS)

    buckets: dict[tuple[int, str, str], dict[str, float | int | str]] = {}

    new_sources = _filter_sources(df_all, list(NEW_HR_SOURCES))
    for _, row in new_sources.iterrows():
        project = clean_text(row.get("project_name"))
        if not project:
            continue

        note = row.get("note")
        date_text = parse_note_value(note, "date")
        report = payroll_report_month_from_date(date_text)
        if report is None:
            continue
        report_year, report_month = report
        if filter_year is not None and report_year != filter_year:
            continue

        for item in MONTHLY_TOTAL_ITEMS:
            amount = parse_note_number(note, item)
            if amount <= 0:
                continue

            key = (report_year, project, item)
            if key not in buckets:
                row_data: dict[str, float | int | str] = {
                    "年度": report_year,
                    "案場": project,
                    "項目": item,
                    "總計": 0.0,
                }
                for label in MONTH_LABELS:
                    row_data[label] = 0.0
                buckets[key] = row_data

            row_data = buckets[key]
            month_col = MONTH_NUM_TO_LABEL[report_month]
            row_data[month_col] = float(row_data.get(month_col) or 0) + amount
            row_data["總計"] = float(row_data.get("總計") or 0) + amount

    if not buckets:
        return pd.DataFrame(columns=MONTHLY_TOTAL_COLS)

    out = pd.DataFrame(list(buckets.values()))
    for col in MONTHLY_TOTAL_COLS:
        if col not in out.columns:
            out[col] = 0.0
    return out.sort_values(["年度", "案場", "項目"], na_position="last")[MONTHLY_TOTAL_COLS]


def build_personal_income_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    buckets: dict[tuple, dict] = {}
    if df_all.empty:
        return pd.DataFrame(columns=PERSONAL_INCOME_COLS)

    new_sources = _filter_sources(df_all, list(NEW_HR_SOURCES))
    for _, row in new_sources.iterrows():
        yr = roc_year_from_value(row.get("roc_year"))
        if yr is None:
            continue
        if filter_year is not None and yr != filter_year:
            continue
        name = clean_text(row.get("employee_name"))
        if not name:
            continue
        note = row.get("note")
        amount = _personal_income_gross(note, row)
        income_tax, business_income, health2 = parse_tax_amounts_from_note(note)
        if amount <= 0 and income_tax == 0 and business_income == 0 and health2 == 0:
            continue
        project = str(row.get("project_name") or "")
        key = (yr, project, name)
        if key not in buckets:
            buckets[key] = {
                "年度": yr,
                "案場": project,
                "姓名": name,
                "金額": 0.0,
                "所得稅": 0.0,
                "執行業務所得": 0.0,
                "二代健保": 0.0,
            }
        entry = buckets[key]
        entry["金額"] += amount
        entry["所得稅"] += income_tax
        entry["執行業務所得"] += business_income
        entry["二代健保"] += health2

    rows: list[dict] = []
    for entry in buckets.values():
        entry["實領金額"] = round(
            entry["金額"] - entry["所得稅"] - entry["執行業務所得"] - entry["二代健保"]
        )
        rows.append(entry)

    if not rows:
        return pd.DataFrame(columns=PERSONAL_INCOME_COLS)
    out = pd.DataFrame(rows)
    return out.sort_values(["年度", "案場", "姓名"], na_position="last")[PERSONAL_INCOME_COLS]
