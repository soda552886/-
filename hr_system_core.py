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

CASE_NOTE_KEYS = ["總銷", "簽約金額", "銷售請款額", "請款額1%", "請款淨額", "營收", "營收(未進帳)"]
HR_NOTE_KEYS = ["勞保", "勞退", "健保", "二代", "薪資", "三節", "獎金", "員工福利", "所得稅", "執行業務所得"]

LEGACY_CASE_SOURCES = {"薪資占比", "總表主表手動", "總表主表調整"}
LEGACY_HR_SOURCES = {"個人所得", "個人所得手動", "薪資獎金統計", "在職年手動", "在職年調整"}
NEW_CASE_SOURCES = {"全案總表", "全案總表手動", "全案總表調整"}
NEW_HR_SOURCES = {"人事成本", "人事成本手動", "人事成本調整"}


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

    if "人事成本" in xls.sheet_names:
        df = pd.read_excel(BytesIO(file_bytes), sheet_name="人事成本", header=0)
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
                total = hr_row_total(note)
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


def _merge_notes_to_row(note_series: pd.Series) -> dict[str, float]:
    merged: dict[str, float] = {}
    for note in note_series:
        for key in CASE_NOTE_KEYS + HR_NOTE_KEYS + ["人事成本", "比例"]:
            merged[key] = merged.get(key, 0.0) + parse_note_number(note, key)
    return merged


def build_case_total_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    sources = _filter_sources(
        df_all,
        ["全案總表", "全案總表手動", "全案總表調整", "薪資占比", "總表主表手動", "總表主表調整"],
    )
    if sources.empty:
        return pd.DataFrame(columns=CASE_TOTAL_COLS)

    hr_agg = build_hr_cost_frame(df_all, filter_year=filter_year)
    hr_lookup: dict[tuple[int, str], float] = {}
    if not hr_agg.empty:
        for _, r in hr_agg.iterrows():
            yr = roc_year_from_value(r.get("年度"))
            if yr is None:
                continue
            hr_lookup[(yr, str(r.get("案場") or ""))] = float(r.get("總計") or 0)

    rows: list[dict] = []
    grouped = sources.groupby(["roc_year", "company_name", "project_name"], dropna=False)
    for (year, company, project), grp in grouped:
        yr = int(year) if pd.notna(year) and to_number(year) >= 100 else None
        if yr is None:
            continue
        if filter_year is not None and yr != filter_year:
            continue
        merged = _merge_notes_to_row(grp["note"])
        if merged.get("總銷", 0) == 0:
            merged["總銷"] = parse_note_number(";".join(grp["note"].astype(str)), "全案總銷")
        sales_request = merged.get("銷售請款額", 0.0)
        request_pct = merged.get("請款額1%", 0.0)
        if request_pct == 0 and sales_request > 0:
            request_pct = calc_request_pct(sales_request)
        hr_cost = merged.get("人事成本", 0.0)
        if hr_cost == 0:
            hr_cost = hr_lookup.get((yr, str(project or "")), 0.0)
        ratio = merged.get("比例", 0.0)
        if ratio == 0:
            ratio = calc_hr_ratio(hr_cost, request_pct)
        rows.append(
            {
                "年度": yr,
                "公司名": str(company or ""),
                "案場": str(project or ""),
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

    if not rows:
        return pd.DataFrame(columns=CASE_TOTAL_COLS)
    out = pd.DataFrame(rows)
    return out.groupby(["年度", "公司名", "案場"], as_index=False).sum(numeric_only=True)[CASE_TOTAL_COLS]


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

    for _, row in df_all.iterrows():
        if is_migrated_note(row.get("note")):
            continue
        if not is_legacy_hr_source(row.get("source_type")):
            continue
        for item in legacy_hr_display_items(row.to_dict()):
            yr = item.get("年度")
            if yr is None:
                continue
            if filter_year is not None and yr != filter_year:
                continue
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
        amount = parse_note_number(note, "薪資") + parse_note_number(note, "獎金")
        if amount <= 0:
            amount = float(row.get("salary") or 0) + float(row.get("bonus") or 0)
        append_yearly(name, yr, amount)

    for _, row in df_all.iterrows():
        if is_migrated_note(row.get("note")):
            continue
        st = str(row.get("source_type") or "")
        if st in {"薪資獎金統計", "在職年手動", "在職年調整"}:
            name = clean_text(row.get("employee_name"))
            if not name:
                continue
            year_amounts = parse_year_amounts(row.get("note"))
            if not year_amounts:
                yr = roc_year_from_value(row.get("roc_year")) or 114
                amt = float(row.get("bonus") or 0) + float(row.get("salary") or 0)
                append_yearly(name, yr, amt)
            else:
                for y_text, amt in year_amounts.items():
                    yr = roc_year_from_value(y_text)
                    if yr is not None:
                        append_yearly(name, yr, amt)
        elif is_legacy_hr_source(st):
            name = clean_text(row.get("employee_name"))
            if not name:
                continue
            for item in legacy_hr_display_items(row.to_dict()):
                append_yearly(name, int(item["年度"]), float(item.get("薪資", 0)) + float(item.get("獎金", 0)))

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


def build_personal_income_frame(df_all: pd.DataFrame, filter_year: int | None = None) -> pd.DataFrame:
    rows: list[dict] = []
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
        salary = parse_note_number(note, "薪資")
        bonus = parse_note_number(note, "獎金")
        amount = salary + bonus
        income_tax = parse_note_number(note, "所得稅")
        business_income = parse_note_number(note, "執行業務所得")
        health2 = parse_note_number(note, "二代")
        net = round(amount - income_tax - business_income - health2)
        rows.append(
            {
                "年度": yr,
                "案場": str(row.get("project_name") or ""),
                "姓名": name,
                "金額": amount,
                "所得稅": income_tax,
                "執行業務所得": business_income,
                "二代健保": health2,
                "實領金額": net,
            }
        )

    for _, row in df_all.iterrows():
        if is_migrated_note(row.get("note")):
            continue
        st = str(row.get("source_type") or "")
        if st in {"個人所得", "個人所得手動"} or st.startswith("個人所得調整"):
            for item in legacy_personal_income_items(row.to_dict()):
                if filter_year is not None and item.get("年度") != filter_year:
                    continue
                rows.append(item)
        elif is_legacy_hr_source(st) and clean_text(row.get("employee_name")):
            for item in legacy_hr_display_items(row.to_dict()):
                yr = item.get("年度")
                if filter_year is not None and yr != filter_year:
                    continue
                amount = float(item.get("薪資", 0)) + float(item.get("獎金", 0))
                note = str(row.get("note") or "")
                income_tax = parse_note_number(note, "所得稅")
                health2 = parse_note_number(note, "二代") or float(row.get("welfare") or 0)
                business = parse_note_number(note, "執行業務所得")
                rows.append(
                    {
                        "年度": yr,
                        "案場": str(row.get("project_name") or item.get("案場") or ""),
                        "姓名": clean_text(row.get("employee_name")),
                        "金額": amount,
                        "所得稅": income_tax,
                        "執行業務所得": business,
                        "二代健保": health2,
                        "實領金額": round(amount - income_tax - business - health2),
                    }
                )

    if not rows:
        return pd.DataFrame(columns=PERSONAL_INCOME_COLS)
    out = pd.DataFrame(rows)
    return out.groupby(["年度", "案場", "姓名"], as_index=False).sum(numeric_only=True)[PERSONAL_INCOME_COLS]
