import re
from io import BytesIO
from typing import List
import os
import socket
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from database import (
    DB_PATH,
    delete_payroll_records,
    delete_records_by_source,
    init_db,
    list_batches,
    list_payroll_records,
    save_import_records,
    update_payroll_record,
)


def format_currency_df(df: pd.DataFrame, cols: List[str]) -> pd.io.formats.style.Styler:
    fmt = {c: "{:,.0f}" for c in cols if c in df.columns}
    numeric_cols = [c for c in cols if c in df.columns]

    def highlight_total_row(row: pd.Series) -> List[str]:
        is_total = str(row.iloc[0]).strip() == "合計"
        if is_total:
            return ["background-color: #f1f5f9; font-weight: 700;" for _ in row]
        return ["" for _ in row]

    styler = (
        df.style.format(fmt)
        .apply(highlight_total_row, axis=1)
        .set_properties(subset=numeric_cols, **{"text-align": "right"})
        .set_properties(subset=[df.columns[0]], **{"text-align": "left"})
    )
    return styler


def to_excel_bytes(
    df: pd.DataFrame,
    sheet_name: str,
    numeric_cols: List[str],
    title: str,
    column_fill_map: dict[str, str] | None = None,
) -> bytes:
    column_fill_map = column_fill_map or {}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        apply_sheet_style(ws, numeric_cols, title, column_fill_map)
    output.seek(0)
    return output.getvalue()


def apply_sheet_style(ws, numeric_cols: List[str], title: str, column_fill_map: dict[str, str]) -> None:
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="E2F0D9")
    border = Border(
        left=Side(style="thin", color="666666"),
        right=Side(style="thin", color="666666"),
        top=Side(style="thin", color="666666"),
        bottom=Side(style="thin", color="666666"),
    )

    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
    title_cell = ws.cell(row=1, column=1, value=title)
    title_cell.font = Font(bold=True, size=14)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=2, column=c)
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    for r in range(3, ws.max_row + 1):
        is_total = str(ws.cell(row=r, column=1).value).strip() == "合計"
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = border
            if is_total:
                cell.fill = total_fill
                cell.font = Font(bold=True)
            else:
                header = str(ws.cell(row=2, column=c).value or "")
                if header in column_fill_map:
                    cell.fill = PatternFill("solid", fgColor=column_fill_map[header])
            header = str(ws.cell(row=2, column=c).value or "")
            if header in numeric_cols:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    preferred_width = {
        "公司名": 20,
        "公司名稱": 20,
        "案場": 28,
        "姓名": 16,
        "分表": 14,
        "project_name": 28,
    }
    for c in range(1, ws.max_column + 1):
        col_letter = get_column_letter(c)
        header = str(ws.cell(row=2, column=c).value or "")
        if header in preferred_width:
            ws.column_dimensions[col_letter].width = preferred_width[header]
            continue
        if header in numeric_cols:
            ws.column_dimensions[col_letter].width = 13
            continue
        if c <= 2:
            ws.column_dimensions[col_letter].width = 20
            continue
        max_len = 0
        for r in range(1, ws.max_row + 1):
            value = ws.cell(row=r, column=c).value
            txt = "" if value is None else str(value)
            max_len = max(max_len, len(txt))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 28)

    ws.freeze_panes = "A3"


def to_excel_multi_sheets(
    sheet_data: dict[str, pd.DataFrame],
    numeric_cols: List[str],
    title_prefix: str,
    column_fill_map: dict[str, str] | None = None,
) -> bytes:
    column_fill_map = column_fill_map or {}
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheet_data.items():
            safe_name = str(sheet_name)[:31] if str(sheet_name).strip() else "Sheet"
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.book[safe_name]
            apply_sheet_style(ws, numeric_cols, f"{title_prefix}（{sheet_name}）", column_fill_map)
    output.seek(0)
    return output.getvalue()


def to_number(value: object) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip().replace(",", "").replace("%", "")
    if text in {"", "-", "nan", "None"}:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def detect_roc_year(text: str) -> int | None:
    match = re.search(r"(\d{3})年", text)
    if match:
        return int(match.group(1))
    return None


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


TOTAL_SUBSHEETS = ["勞健退", "薪資", "三節", "獎金", "餐費", "員工福利"]
TOTAL_MAIN_ITEM_COLS = ["勞健退", "薪資", "三節", "獎金", "餐費", "員工福利"]
TOTAL_MAIN_ALL_COLS = ["公司名", "案場", *TOTAL_MAIN_ITEM_COLS, "人事成本總計", "原始總計"]
PROJECT_SALES_COLS = ["案量", "銷售金額", "銷售佣金", "實收佣金", "廣告費", "服務費"]
PROJECT_HR_COLS = ["薪資", "勞健保", "退休金", "獎金"]
PROJECT_SUMMARY_COLS = ["公司名", "案名", *PROJECT_SALES_COLS, *PROJECT_HR_COLS, "人事成本合計"]
PROJECT_SUMMARY_ALL_COLS = PROJECT_SUMMARY_COLS
COMPANY_OPTIONS = ["得意佳", "匯鴻", "鴻源", "寶得", "得威", "賦鼎"]
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
INCOME_TYPE_OPTIONS = ["執行業務所得", "四倍獎金累計"]


def parse_manual_category(note: object) -> str:
    text = "" if pd.isna(note) else str(note)
    match = re.search(r"category:([^;]+)", text)
    if match:
        return match.group(1).strip()
    return ""


def parse_note_value(note: object, key: str) -> str:
    text = "" if pd.isna(note) else str(note)
    match = re.search(rf"{re.escape(key)}:([^;]+)", text)
    if match:
        return match.group(1).strip()
    return ""


def parse_note_number(note: object, key: str) -> float:
    return to_number(parse_note_value(note, key))


def append_note_parts(parts: list[str]) -> str:
    return ";".join([p for p in parts if str(p).strip()])


def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"


def calc_income_deductions(amount: float, income_type: str) -> tuple[float, float, float]:
    gross = float(amount or 0)
    if income_type == "執行業務所得":
        tax = round(gross * 0.10)
        health = 0.0
    else:
        tax = round(gross * 0.05)
        health = round(gross * 0.0211)
    net = round(gross - tax - health)
    return tax, health, net


def to_main_like_columns(df: pd.DataFrame, category: str) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=TOTAL_MAIN_ALL_COLS)

    out = pd.DataFrame()
    out["公司名"] = df["公司名"]
    out["案場"] = df["案場"]
    for item in TOTAL_MAIN_ITEM_COLS:
        out[item] = df["金額"] if category == item else 0.0
    out["人事成本總計"] = out[TOTAL_MAIN_ITEM_COLS].sum(axis=1)
    out["原始總計"] = out["人事成本總計"]
    return out


def build_main_total_frame(df_all: pd.DataFrame) -> pd.DataFrame:
    sr = df_all[df_all["source_type"] == "薪資占比"].copy()
    main_manual = df_all[df_all["source_type"] == "總表主表手動"].copy()
    main_adjust = df_all[df_all["source_type"] == "總表主表調整"].copy()

    rows = []
    for _, row in sr.iterrows():
        rows.append(
            {
                "公司名": str(row.get("company_name") or "").strip(),
                "案場": str(row.get("project_name") or "").strip(),
                "勞健退": 0.0,
                "薪資": float(row.get("salary") or 0),
                "三節": 0.0,
                "獎金": float(row.get("bonus") or 0),
                "餐費": 0.0,
                "員工福利": float(row.get("welfare") or 0),
                "原始總計": float(row.get("total_income") or 0),
            }
        )
    for _, row in main_manual.iterrows():
        salary = float(row.get("salary") or 0)
        bonus = float(row.get("bonus") or 0)
        welfare = float(row.get("welfare") or 0)
        original = float(row.get("total_income") or 0)
        labor = parse_note_number(row.get("note"), "勞健退")
        festival = parse_note_number(row.get("note"), "三節")
        meal = parse_note_number(row.get("note"), "餐費")
        if original <= 0:
            original = labor + salary + festival + bonus + meal + welfare
        rows.append(
            {
                "公司名": str(row.get("company_name") or "").strip(),
                "案場": str(row.get("project_name") or "").strip(),
                "勞健退": labor,
                "薪資": salary,
                "三節": festival,
                "獎金": bonus,
                "餐費": meal,
                "員工福利": welfare,
                "原始總計": original,
            }
        )
    for _, row in main_adjust.iterrows():
        rows.append(
            {
                "公司名": str(row.get("company_name") or "").strip(),
                "案場": str(row.get("project_name") or "").strip(),
                "勞健退": parse_note_number(row.get("note"), "勞健退"),
                "薪資": float(row.get("salary") or 0),
                "三節": parse_note_number(row.get("note"), "三節"),
                "獎金": float(row.get("bonus") or 0),
                "餐費": parse_note_number(row.get("note"), "餐費"),
                "員工福利": float(row.get("welfare") or 0),
                "原始總計": float(row.get("total_income") or 0),
            }
        )

    grouped = pd.DataFrame(rows)
    if grouped.empty:
        return pd.DataFrame(columns=TOTAL_MAIN_ALL_COLS)
    grouped = grouped.groupby(["公司名", "案場"], as_index=False)[[*TOTAL_MAIN_ITEM_COLS, "原始總計"]].sum()
    grouped["人事成本總計"] = grouped[TOTAL_MAIN_ITEM_COLS].sum(axis=1)
    return grouped[TOTAL_MAIN_ALL_COLS]


def build_total_subsheet_frames(df_all: pd.DataFrame) -> dict[str, pd.DataFrame]:
    bucket: dict[str, List[dict]] = {k: [] for k in TOTAL_SUBSHEETS}

    # 從薪資占比匯入資料拆成既有可對應分表
    sr = df_all[df_all["source_type"] == "薪資占比"].copy()
    for _, row in sr.iterrows():
        company = str(row.get("company_name") or "").strip()
        project = str(row.get("project_name") or "").strip()
        if company:
            if float(row.get("salary") or 0) > 0:
                bucket["薪資"].append({"公司名": company, "案場": project, "金額": float(row.get("salary") or 0), "備註": "來源: 薪資占比"})
            if float(row.get("bonus") or 0) > 0:
                bucket["獎金"].append({"公司名": company, "案場": project, "金額": float(row.get("bonus") or 0), "備註": "來源: 薪資占比"})
            if float(row.get("welfare") or 0) > 0:
                bucket["員工福利"].append({"公司名": company, "案場": project, "金額": float(row.get("welfare") or 0), "備註": "來源: 薪資占比"})

    # 手動輸入資料
    manual = df_all[
        df_all["source_type"].isin(["總表分表手動", "總表分表匯入"])
        | df_all["source_type"].fillna("").astype(str).str.startswith("總表分表調整_")
    ].copy()
    for _, row in manual.iterrows():
        category = parse_manual_category(row.get("note"))
        if category not in bucket:
            continue
        bucket[category].append(
            {
                "公司名": str(row.get("company_name") or "").strip(),
                "案場": str(row.get("project_name") or "").strip(),
                "金額": float(row.get("total_income") or 0),
                "備註": str(row.get("note") or "").replace(f"category:{category};", "").strip(),
            }
        )

    frames: dict[str, pd.DataFrame] = {}
    for category, rows in bucket.items():
        if rows:
            df = pd.DataFrame(rows)
            df = df.groupby(["公司名", "案場"], as_index=False)["金額"].sum()
            total_row = pd.DataFrame([{"公司名": "合計", "案場": "", "金額": df["金額"].sum()}])
            frames[category] = pd.concat([df, total_row], ignore_index=True)
        else:
            frames[category] = pd.DataFrame(columns=["公司名", "案場", "金額"])
    return frames


def subsheet_amount_map(df_all: pd.DataFrame) -> dict[tuple[str, str], dict[str, float]]:
    frames = build_total_subsheet_frames(df_all)
    out: dict[tuple[str, str], dict[str, float]] = {}
    for category, df in frames.items():
        if df.empty:
            continue
        for _, row in df.iterrows():
            company = str(row.get("公司名") or "").strip()
            project = str(row.get("案場") or "").strip()
            if not company or company == "合計" or not project:
                continue
            key = (company, project)
            out.setdefault(key, {})
            out[key][category] = out[key].get(category, 0.0) + float(row.get("金額") or 0)
    return out


def build_project_summary_frame(df_all: pd.DataFrame) -> pd.DataFrame:
    sources = df_all[df_all["source_type"].isin(["薪資占比", "總表主表手動", "總表主表調整"])].copy()
    sub_map = subsheet_amount_map(df_all)
    rows: list[dict] = []

    def append_row(company: str, project: str, note: object, salary: float, bonus: float, welfare: float) -> None:
        if not company or not project:
            return
        key = (company, project)
        sub = sub_map.get(key, {})
        hr_salary = sub.get("薪資", 0.0) or salary
        hr_labor = sub.get("勞健退", 0.0)
        hr_pension = 0.0
        hr_bonus = sub.get("獎金", 0.0) or bonus
        if hr_salary <= 0 and salary > 0:
            hr_salary = salary
        if hr_bonus <= 0 and bonus > 0:
            hr_bonus = bonus
        if hr_labor <= 0:
            hr_labor = parse_note_number(note, "勞健退")
        rows.append(
            {
                "公司名": company,
                "案名": project,
                "案量": parse_note_number(note, "全案總銷"),
                "銷售金額": parse_note_number(note, "簽約金額"),
                "銷售佣金": parse_note_number(note, "總銷1%"),
                "實收佣金": parse_note_number(note, "簽約金額1%"),
                "廣告費": parse_note_number(note, "營業收入"),
                "服務費": parse_note_number(note, "請款淨額"),
                "薪資": hr_salary,
                "勞健保": hr_labor,
                "退休金": hr_pension,
                "獎金": hr_bonus,
            }
        )

    for _, row in sources.iterrows():
        company = str(row.get("company_name") or "").strip()
        project = str(row.get("project_name") or "").strip()
        append_row(
            company,
            project,
            row.get("note"),
            float(row.get("salary") or 0),
            float(row.get("bonus") or 0),
            float(row.get("welfare") or 0),
        )

    if not rows:
        return pd.DataFrame(columns=PROJECT_SUMMARY_ALL_COLS)

    grouped = pd.DataFrame(rows).groupby(["公司名", "案名"], as_index=False).sum(numeric_only=True)
    grouped["人事成本合計"] = grouped[PROJECT_HR_COLS].sum(axis=1)
    return grouped[PROJECT_SUMMARY_ALL_COLS]


def parse_total_subsheet_workbook(file_path: str) -> List[dict]:
    xls = pd.ExcelFile(file_path)
    records: List[dict] = []

    for category in TOTAL_SUBSHEETS:
        if category not in xls.sheet_names:
            continue
        df = pd.read_excel(file_path, sheet_name=category, header=3)
        if df.empty or df.shape[1] < 3:
            continue

        col_company = df.columns[0]
        col_project = df.columns[1]
        month_cols = [c for c in df.columns[2:] if isinstance(c, str) and "/" in c]
        if not month_cols:
            month_cols = list(df.columns[2:])

        current_company = ""
        for _, row in df.iterrows():
            company_raw = clean_text(row.get(col_company))
            project = clean_text(row.get(col_project))
            if company_raw:
                current_company = company_raw.replace("\n", "")
            company = current_company
            if not company or not project:
                continue

            values = [to_number(row.get(c)) for c in month_cols]
            amount = float(sum(values))
            if amount <= 0:
                continue

            records.append(
                {
                    "sheet_name": "全案總表分表",
                    "employee_name": "未指定",
                    "company_name": company,
                    "project_name": project,
                    "roc_year": None,
                    "salary": 0.0,
                    "bonus": 0.0,
                    "welfare": 0.0,
                    "total_income": amount,
                    "note": f"category:{category};來源:薪資占比檔分表",
                }
            )
    return records


def upsert_adjustment_records(source_type: str, records: List[dict]) -> None:
    delete_records_by_source(source_type)
    if records:
        save_import_records(source_type, "inline_edit", records)


def parse_year_amounts(note: object) -> dict[str, float]:
    text = "" if pd.isna(note) else str(note)
    pairs = re.findall(r"(\d{2,4})\s*:\s*([-\d.]+)", text)
    out: dict[str, float] = {}
    for y, v in pairs:
        out[y] = float(pd.to_numeric(v, errors="coerce") or 0.0)
    return out


def parse_personal_income_detail_sheet(df: pd.DataFrame, sheet_name: str) -> List[dict]:
    records: List[dict] = []
    n_rows, n_cols = df.shape
    roc_year = detect_roc_year(sheet_name) or 114

    for i in range(n_rows):
        row = [clean_text(v) for v in df.iloc[i].tolist()]
        for start in range(max(0, n_cols - 2)):
            employee = row[start]
            if not employee:
                continue
            if start + 2 >= n_cols:
                continue
            if "案場" not in row[start + 1] or "金額" not in row[start + 2]:
                continue

            employee_name = employee.replace("*", "").replace("總計", "").strip()
            if not employee_name or employee_name in {"姓名", "投保級距", "非投保單位"}:
                continue

            # 依每個人員區塊的表頭，找出目標欄位位置
            header_window = row[start : min(start + 14, n_cols)]
            idx_project = 1
            idx_amount = 2
            idx_tax = next((k for k, v in enumerate(header_window) if "所得稅" in v or "應扣稅額" in v), None)
            idx_health = next((k for k, v in enumerate(header_window) if "二代健保" in v), None)
            idx_net = next((k for k, v in enumerate(header_window) if "實領" in v), None)
            idx_withhold_tax = next((k for k, v in enumerate(header_window) if "應扣所得" in v or "應扣所得稅" in v), None)
            idx_withhold_health = next((k for k, v in enumerate(header_window) if "應扣二代" in v), None)

            for r in range(i + 1, n_rows):
                cur = [clean_text(v) for v in df.iloc[r].tolist()]
                first = cur[start] if start < n_cols else ""
                project = cur[start + idx_project] if start + idx_project < n_cols else ""
                amount = to_number(cur[start + idx_amount] if start + idx_amount < n_cols else 0)

                # 遇到總計列或下一段人員表頭，結束此人員區塊
                if "總計" in first:
                    break
                if first and project == "案場":
                    break

                if not project or project in {"案場", "-"}:
                    continue
                if amount <= 0:
                    continue

                tax = to_number(cur[start + idx_tax] if idx_tax is not None and start + idx_tax < n_cols else 0)
                health = to_number(cur[start + idx_health] if idx_health is not None and start + idx_health < n_cols else 0)
                net = to_number(cur[start + idx_net] if idx_net is not None and start + idx_net < n_cols else amount)
                withhold_tax = to_number(
                    cur[start + idx_withhold_tax] if idx_withhold_tax is not None and start + idx_withhold_tax < n_cols else 0
                )
                withhold_health = to_number(
                    cur[start + idx_withhold_health] if idx_withhold_health is not None and start + idx_withhold_health < n_cols else 0
                )

                records.append(
                    {
                        "sheet_name": sheet_name,
                        "employee_name": employee_name,
                        "company_name": sheet_name,
                        "project_name": project,
                        "roc_year": roc_year,
                        "salary": amount,
                        "bonus": tax,
                        "welfare": health,
                        "total_income": net,
                        "note": f"應扣所得:{withhold_tax}, 應扣二代:{withhold_health}",
                    }
                )
    return records


def parse_salary_ratio_workbook(file_bytes: bytes) -> List[dict]:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    target_sheet = "全案_總表" if "全案_總表" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=3)
    df.columns = [str(c).strip().replace("\n", "") for c in df.columns]

    col_company = next((c for c in df.columns if "公司" in c), None)
    col_project = next((c for c in df.columns if "案場" in c or c == "案名"), None)
    col_total_sales = next((c for c in df.columns if "全案總銷" in c), None)
    col_booking = next((c for c in df.columns if c == "記帳"), None)
    col_contract = next((c for c in df.columns if "簽約金額" in c and "%" not in c), None)
    col_net_request = next((c for c in df.columns if "請款淨額" in c), None)
    col_revenue = next((c for c in df.columns if c == "營業收入"), None)
    col_sales_pct = next((c for c in df.columns if c == "總銷1%"), None)
    col_contract_pct = next((c for c in df.columns if "簽約金額1%" in c), None)
    col_salary = next((c for c in df.columns if "薪資" in c and "含" in c), None)
    if not col_salary:
        col_salary = next((c for c in df.columns if c == "薪資" or "薪資" in c), None)
    col_bonus = next((c for c in df.columns if "獎金" in c), None)
    col_welfare = next((c for c in df.columns if "福利" in c), None)
    col_total = next((c for c in df.columns if c == "總計"), None)

    if not col_company or not col_project:
        raise ValueError("無法辨識『全案_總表』的公司/案場欄位。")

    records: List[dict] = []
    for _, row in df.iterrows():
        company = str(row.get(col_company, "")).strip()
        project = str(row.get(col_project, "")).strip()
        if not company or company in {"nan", "合計"}:
            continue
        records.append(
            {
                "sheet_name": target_sheet,
                "employee_name": None,
                "company_name": company,
                "project_name": project,
                "roc_year": None,
                "salary": to_number(row.get(col_salary)),
                "bonus": to_number(row.get(col_bonus)),
                "welfare": to_number(row.get(col_welfare)),
                "total_income": to_number(row.get(col_total)),
                "note": append_note_parts(
                    [
                        "來源: 薪資占比檔",
                        f"全案總銷:{to_number(row.get(col_total_sales))}",
                        f"記帳:{to_number(row.get(col_booking))}",
                        f"簽約金額:{to_number(row.get(col_contract))}",
                        f"請款淨額:{to_number(row.get(col_net_request))}",
                        f"營業收入:{to_number(row.get(col_revenue))}",
                        f"總銷1%:{to_number(row.get(col_sales_pct))}",
                        f"簽約金額1%:{to_number(row.get(col_contract_pct))}",
                    ]
                ),
            }
        )
    return records


def parse_bonus_stat_workbook(file_bytes: bytes) -> List[dict]:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    target_sheet = "在職年統計" if "在職年統計" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=target_sheet, header=0)
    df.columns = [str(c).strip() for c in df.columns]

    records: List[dict] = []
    for i in [0, 1]:
        suffix = "" if i == 0 else ".1"
        col_name = f"姓名{suffix}" if f"姓名{suffix}" in df.columns else ("姓名" if i == 0 else None)
        if not col_name:
            continue
        col_113 = f"113年{suffix}" if f"113年{suffix}" in df.columns else None
        col_114 = f"114年{suffix}" if f"114年{suffix}" in df.columns else None
        col_115 = f"115年{suffix}" if f"115年{suffix}" in df.columns else None
        col_sum = f"合計{suffix}" if f"合計{suffix}" in df.columns else None

        for _, row in df.iterrows():
            name = str(row.get(col_name, "")).strip()
            if not name or name in {"nan", "姓名"}:
                continue
            records.append(
                {
                    "sheet_name": target_sheet,
                    "employee_name": name,
                    "company_name": None,
                    "project_name": None,
                    "roc_year": 114,
                    "salary": 0.0,
                    "bonus": to_number(row.get(col_114)),
                    "welfare": 0.0,
                    "total_income": to_number(row.get(col_sum)),
                    "note": f"113:{to_number(row.get(col_113))}, 114:{to_number(row.get(col_114))}, 115:{to_number(row.get(col_115))}",
                }
            )
    return records


def parse_personal_income_workbook(file_bytes: bytes) -> List[dict]:
    xls = pd.ExcelFile(BytesIO(file_bytes))
    records: List[dict] = []

    for sheet in xls.sheet_names:
        # 跳過彙總/說明頁，主抓每人每案場明細頁
        if str(sheet).startswith("總計-") or str(sheet).startswith("薪資-") or str(sheet) == "扣繳":
            continue

        df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet, header=None)
        if df.empty:
            continue
        records.extend(parse_personal_income_detail_sheet(df, sheet))
    return records


st.set_page_config(page_title="薪資報表匯入管理系統", layout="wide")
init_db()

st.title("薪資報表匯入管理系統")
st.caption("依你提供的三份 Excel 結構設計：專用匯入、批次存檔、查詢瀏覽。")
local_ip = get_local_ip()
with st.expander("跨裝置開啟網站"):
    st.write("在同一個 Wi-Fi/區網下，請用以下方式啟動：")
    st.code("streamlit run app.py --server.address 0.0.0.0 --server.port 8501")
    st.write(f"其他裝置可開：`http://{local_ip}:8501`（請確認 Windows 防火牆允許 8501 連線）")

tab_import, tab_report, tab_query, tab_batches = st.tabs(["匯入資料", "報表呈現", "資料查詢", "匯入紀錄"])

with tab_import:
    st.subheader("1) 薪資占比1150131.xlsx")
    file1 = st.file_uploader("上傳薪資占比檔", type=["xlsx", "xls"], key="upload_salary_ratio")
    replace_salary_ratio = st.checkbox("匯入前刪除舊的「薪資占比」資料", key="replace_salary_ratio")
    if file1 is not None and st.button("匯入薪資占比", key="import_salary_ratio"):
        try:
            if replace_salary_ratio:
                deleted = delete_records_by_source("薪資占比")
                st.info(f"已先刪除舊資料 {deleted} 筆。")
            records = parse_salary_ratio_workbook(file1.getvalue())
            inserted = save_import_records("薪資占比", file1.name, records)
            st.success(f"匯入成功，共 {inserted} 筆。")
            if records:
                st.dataframe(pd.DataFrame(records).head(20), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"匯入失敗：{exc}")

    st.subheader("2) 114薪資獎金統計表1150131(更.xlsx")
    file2 = st.file_uploader("上傳薪資獎金統計檔", type=["xlsx", "xls"], key="upload_bonus_stat")
    replace_bonus_stat = st.checkbox("匯入前刪除舊的「薪資獎金統計」資料", key="replace_bonus_stat")
    if file2 is not None and st.button("匯入薪資獎金統計", key="import_bonus_stat"):
        try:
            if replace_bonus_stat:
                deleted = delete_records_by_source("薪資獎金統計")
                st.info(f"已先刪除舊資料 {deleted} 筆。")
            records = parse_bonus_stat_workbook(file2.getvalue())
            inserted = save_import_records("薪資獎金統計", file2.name, records)
            st.success(f"匯入成功，共 {inserted} 筆。")
            if records:
                st.dataframe(pd.DataFrame(records).head(20), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"匯入失敗：{exc}")

    st.subheader("3) 114年_個人所得.xlsx")
    file3 = st.file_uploader("上傳個人所得檔", type=["xlsx", "xls"], key="upload_personal_income")
    replace_personal_income = st.checkbox("匯入前刪除舊的「個人所得」資料", key="replace_personal_income")
    if file3 is not None and st.button("匯入個人所得", key="import_personal_income"):
        try:
            if replace_personal_income:
                deleted = delete_records_by_source("個人所得")
                st.info(f"已先刪除舊資料 {deleted} 筆。")
            records = parse_personal_income_workbook(file3.getvalue())
            inserted = save_import_records("個人所得", file3.name, records)
            st.success(f"匯入成功，共 {inserted} 筆。")
            if records:
                st.dataframe(pd.DataFrame(records).head(20), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(f"匯入失敗：{exc}")

with tab_report:
    st.subheader("依原始檔風格呈現報表")
    records = list_payroll_records(limit=100000)
    if not records:
        st.info("目前沒有資料，請先到「匯入資料」上傳檔案。")
    else:
        df_all = pd.DataFrame([dict(r) for r in records])
        report_view = st.selectbox(
            "選擇報表樣式",
            ["專案總表（薪資占比）", "全案總表（六分表）", "在職年統計（薪資獎金）", "個人所得明細"],
        )

        if report_view == "專案總表（薪資占比）":
            grouped = build_project_summary_frame(df_all)
            if grouped.empty:
                st.warning("尚未匯入「薪資占比」資料。請到「匯入資料」重新匯入，總銷等欄位才會出現。")
            else:
                visible_total_cols = st.multiselect(
                    "專案總表顯示項目（不需要的欄位可取消勾選）",
                    options=PROJECT_SUMMARY_ALL_COLS,
                    default=PROJECT_SUMMARY_ALL_COLS,
                    key="visible_project_cols",
                )
                numeric_cols = [c for c in PROJECT_SALES_COLS + PROJECT_HR_COLS + ["人事成本合計"] if c in grouped.columns]
                total_row = pd.DataFrame(
                    [
                        {
                            "公司名": "合計",
                            "案名": "",
                            **{col: grouped[col].sum() for col in numeric_cols},
                        }
                    ]
                )
                display_df = pd.concat([grouped, total_row], ignore_index=True)
                shown_cols = visible_total_cols if visible_total_cols else PROJECT_SUMMARY_ALL_COLS
                st.dataframe(
                    format_currency_df(
                        display_df[shown_cols],
                        [c for c in numeric_cols if c in shown_cols],
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
                st.download_button(
                    "匯出專案總表 Excel",
                    data=to_excel_bytes(
                        display_df[shown_cols],
                        "專案總表",
                        numeric_cols=[c for c in numeric_cols if c in shown_cols],
                        title="專案總表（薪資占比）",
                        column_fill_map={
                            **{c: "FCE4D6" for c in PROJECT_SALES_COLS},
                            **{c: "E2F0D9" for c in PROJECT_HR_COLS},
                            "人事成本合計": "FFF2CC",
                        },
                    ),
                    file_name="專案總表_匯出.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_project_summary_report",
                )
                st.caption("若總銷欄位為 0，請重新匯入薪資占比檔以更新完整欄位。")

        elif report_view == "全案總表（六分表）":
            grouped = build_main_total_frame(df_all)
            if grouped.empty:
                st.warning("尚未匯入「薪資占比」資料。")
            else:
                visible_total_cols = st.multiselect(
                    "總表顯示項目",
                    options=TOTAL_MAIN_ALL_COLS,
                    default=TOTAL_MAIN_ALL_COLS,
                    key="visible_total_cols",
                )
                total_row = pd.DataFrame(
                    [
                        {
                            "公司名": "合計",
                            "案場": "",
                            "勞健退": grouped["勞健退"].sum(),
                            "薪資": grouped["薪資"].sum(),
                            "三節": grouped["三節"].sum(),
                            "獎金": grouped["獎金"].sum(),
                            "餐費": grouped["餐費"].sum(),
                            "員工福利": grouped["員工福利"].sum(),
                            "人事成本總計": grouped["人事成本總計"].sum(),
                            "原始總計": grouped["原始總計"].sum(),
                        }
                    ]
                )
                display_df = pd.concat([grouped, total_row], ignore_index=True)
                sub_frames = build_total_subsheet_frames(df_all)
                editable_main = display_df.iloc[:-1].copy()
                shown_cols = visible_total_cols if visible_total_cols else TOTAL_MAIN_ALL_COLS
                edited_visible = st.data_editor(
                    editable_main[shown_cols],
                    use_container_width=True,
                    hide_index=True,
                    key="main_report_inline_edit",
                    disabled=[],
                )
                edited_main = editable_main.copy()
                for col in shown_cols:
                    edited_main[col] = edited_visible[col]
                if st.button("儲存這個主表格的修改", key="save_main_report_inline"):
                    # 以「調整量」方式儲存，確保可直接覆寫目前畫面數字
                    base = editable_main.copy()
                    now = edited_main.copy()
                    for col in TOTAL_MAIN_ITEM_COLS + ["原始總計", "人事成本總計"]:
                        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0.0)
                        now[col] = pd.to_numeric(now[col], errors="coerce").fillna(0.0)

                    delete_records_by_source("總表主表調整")
                    adjust_records = []
                    for i in range(len(now)):
                        old_company = str(base.iloc[i]["公司名"])
                        old_project = str(base.iloc[i]["案場"])
                        new_company = str(now.iloc[i]["公司名"])
                        new_project = str(now.iloc[i]["案場"])
                        old_vals = {col: float(base.iloc[i][col]) for col in TOTAL_MAIN_ITEM_COLS + ["原始總計"]}
                        new_vals = {col: float(now.iloc[i][col]) for col in TOTAL_MAIN_ITEM_COLS + ["原始總計"]}

                        if old_company == new_company and old_project == new_project:
                            deltas = {col: new_vals[col] - old_vals[col] for col in TOTAL_MAIN_ITEM_COLS + ["原始總計"]}
                            if all(abs(v) < 1e-6 for v in deltas.values()):
                                continue
                            adjust_records.append(
                                {
                                    "sheet_name": "全案總表",
                                    "employee_name": "未指定",
                                    "company_name": new_company,
                                    "project_name": new_project,
                                    "roc_year": None,
                                    "salary": deltas["薪資"],
                                    "bonus": deltas["獎金"],
                                    "welfare": deltas["員工福利"],
                                    "total_income": deltas["原始總計"],
                                    "note": append_note_parts([
                                        "inline_main_adjustment",
                                        f"勞健退:{deltas['勞健退']}",
                                        f"三節:{deltas['三節']}",
                                        f"餐費:{deltas['餐費']}",
                                    ]),
                                }
                            )
                        else:
                            if any(abs(v) > 1e-6 for v in old_vals.values()):
                                adjust_records.append(
                                    {
                                        "sheet_name": "全案總表",
                                        "employee_name": "未指定",
                                        "company_name": old_company,
                                        "project_name": old_project,
                                        "roc_year": None,
                                        "salary": -old_vals["薪資"],
                                        "bonus": -old_vals["獎金"],
                                        "welfare": -old_vals["員工福利"],
                                        "total_income": -old_vals["原始總計"],
                                        "note": append_note_parts([
                                            "inline_main_adjustment",
                                            f"勞健退:{-old_vals['勞健退']}",
                                            f"三節:{-old_vals['三節']}",
                                            f"餐費:{-old_vals['餐費']}",
                                        ]),
                                    }
                                )
                            if any(abs(v) > 1e-6 for v in new_vals.values()):
                                adjust_records.append(
                                    {
                                        "sheet_name": "全案總表",
                                        "employee_name": "未指定",
                                        "company_name": new_company,
                                        "project_name": new_project,
                                        "roc_year": None,
                                        "salary": new_vals["薪資"],
                                        "bonus": new_vals["獎金"],
                                        "welfare": new_vals["員工福利"],
                                        "total_income": new_vals["原始總計"],
                                        "note": append_note_parts([
                                            "inline_main_adjustment",
                                            f"勞健退:{new_vals['勞健退']}",
                                            f"三節:{new_vals['三節']}",
                                            f"餐費:{new_vals['餐費']}",
                                        ]),
                                    }
                                )
                    if adjust_records:
                        save_import_records("總表主表調整", "inline_edit", adjust_records)
                    st.success("已儲存主表格修改。")
                    st.rerun()
                st.download_button(
                    "匯出總表+六分表 Excel（多工作表）",
                    data=to_excel_multi_sheets(
                        sheet_data={
                            "主總表": display_df,
                            "勞健退": to_main_like_columns(sub_frames["勞健退"], "勞健退"),
                            "薪資": to_main_like_columns(sub_frames["薪資"], "薪資"),
                            "三節": to_main_like_columns(sub_frames["三節"], "三節"),
                            "獎金": to_main_like_columns(sub_frames["獎金"], "獎金"),
                            "餐費": to_main_like_columns(sub_frames["餐費"], "餐費"),
                            "員工福利": to_main_like_columns(sub_frames["員工福利"], "員工福利"),
                        },
                        numeric_cols=TOTAL_MAIN_ITEM_COLS + ["人事成本總計", "原始總計"],
                        title_prefix="全案總表",
                        column_fill_map={
                            "勞健退": "FFF2CC",
                            "薪資": "FFF2CC",
                            "三節": "FFF2CC",
                            "獎金": "FFF2CC",
                            "餐費": "FFF2CC",
                            "員工福利": "FFF2CC",
                            "人事成本總計": "FCE4D6",
                            "原始總計": "E2F0D9",
                        },
                    ),
                    file_name="全案總表_含六分表_匯出.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_salary_ratio_report",
                )
                st.caption("此表模擬你原檔中的『全案_總表』主體資料區。")

                st.markdown("### 手動新增資料")
                add_tabs = st.tabs(["新增總表", "新增個人資料"])
                with add_tabs[0]:
                    company_options = COMPANY_OPTIONS + ["自訂輸入"]
                    with st.form("manual_add_main_form", clear_on_submit=True):
                        c1, c2, c3 = st.columns(3)
                        with c1:
                            manual_company_select = st.selectbox("公司名", company_options, key="add_main_company_select")
                        with c2:
                            manual_project_select = st.selectbox(
                                "案名",
                                PROJECT_OPTIONS,
                                key="add_main_project_select",
                            )
                        with c3:
                            manual_date = st.date_input("日期", key="add_main_date")

                        if manual_company_select == "自訂輸入":
                            manual_company = st.text_input("自訂公司名", key="add_main_company_custom")
                        else:
                            manual_company = manual_company_select
                        manual_project = manual_project_select

                        c4, c5, c6 = st.columns(3)
                        with c4:
                            manual_labor = st.number_input("勞健退", min_value=0.0, step=1000.0, format="%.0f", key="add_main_labor")
                        with c5:
                            manual_salary = st.number_input("薪資", min_value=0.0, step=1000.0, format="%.0f", key="add_main_salary")
                        with c6:
                            manual_festival = st.number_input("三節", min_value=0.0, step=1000.0, format="%.0f", key="add_main_festival")
                        c7, c8, c9 = st.columns(3)
                        with c7:
                            manual_bonus = st.number_input("獎金", min_value=0.0, step=1000.0, format="%.0f", key="add_main_bonus")
                        with c8:
                            manual_meal = st.number_input("餐費", min_value=0.0, step=1000.0, format="%.0f", key="add_main_meal")
                        with c9:
                            manual_welfare = st.number_input("員工福利", min_value=0.0, step=1000.0, format="%.0f", key="add_main_welfare")
                        c10, c11 = st.columns(2)
                        with c10:
                            manual_total = st.number_input("原始總計（留 0 自動加總）", min_value=0.0, step=1000.0, format="%.0f", key="add_main_total")
                        with c11:
                            manual_note = st.text_input("備註（可空白）", key="add_main_note")
                        submit_main = st.form_submit_button("新增總表資料")
                    if submit_main:
                        if not str(manual_company).strip() or not str(manual_project).strip():
                            st.error("請輸入公司名與案名。")
                        else:
                            save_import_records(
                                "總表主表手動",
                                "manual_input_main",
                                [{
                                    "sheet_name": "全案總表",
                                    "employee_name": "未指定",
                                    "company_name": str(manual_company).strip(),
                                    "project_name": str(manual_project).strip(),
                                    "roc_year": None,
                                    "salary": float(manual_salary),
                                    "bonus": float(manual_bonus),
                                    "welfare": float(manual_welfare),
                                    "total_income": float(manual_total),
                                    "note": append_note_parts([
                                        "manual_main",
                                        f"date:{manual_date.isoformat()}",
                                        f"勞健退:{float(manual_labor)}",
                                        f"三節:{float(manual_festival)}",
                                        f"餐費:{float(manual_meal)}",
                                        manual_note.strip(),
                                    ]),
                                }],
                            )
                            st.success("已新增總表資料。")
                            st.rerun()
                with add_tabs[1]:
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        person_company = st.selectbox("公司名", COMPANY_OPTIONS, key="add_person_company")
                    with c2:
                        income_name = st.text_input("姓名", key="add_person_name")
                    with c3:
                        income_project_select = st.selectbox(
                            "案名",
                            PROJECT_OPTIONS,
                            key="add_person_project_select",
                        )
                    with c4:
                        person_date = st.date_input("日期", key="add_person_date")
                    income_project = income_project_select

                    c5, c6, c7, c8 = st.columns(4)
                    with c5:
                        income_total = st.number_input("應發額", min_value=0.0, step=1000.0, format="%.0f", key="add_income_total")
                    with c6:
                        income_type = st.selectbox("所得類型", INCOME_TYPE_OPTIONS, key="add_income_type")
                    with c7:
                        auto_calc_income = st.checkbox("自動帶入所得/二代/合計", value=True, key="add_income_auto")
                    with c8:
                        sync_yearly = st.checkbox("同步寫入在職年統計", value=True, key="add_sync_yearly")

                    calc_tax, calc_health, calc_net = calc_income_deductions(income_total, income_type)
                    p1, p2, p3 = st.columns(3)
                    if auto_calc_income:
                        p1.metric("所得", f"{calc_tax:,.0f}")
                        p2.metric("二代健保費", f"{calc_health:,.0f}")
                        p3.metric("合計（實領）", f"{calc_net:,.0f}")
                        income_tax = calc_tax
                        income_health = calc_health
                        income_net = calc_net
                    else:
                        with p1:
                            income_tax = st.number_input("所得", min_value=0.0, step=100.0, format="%.0f", key="add_income_tax")
                        with p2:
                            income_health = st.number_input("二代健保費", min_value=0.0, step=100.0, format="%.0f", key="add_income_health")
                        with p3:
                            income_net = st.number_input("合計（實領）", min_value=0.0, step=1000.0, format="%.0f", key="add_income_net")

                    with st.form("manual_add_person_form", clear_on_submit=True):
                        c11, c12, c13 = st.columns(3)
                        with c11:
                            income_wtax = st.number_input("應扣所得", min_value=0.0, step=100.0, format="%.0f", key="add_income_wtax")
                        with c12:
                            income_whealth = st.number_input("應扣二代", min_value=0.0, step=100.0, format="%.0f", key="add_income_whealth")
                        with c13:
                            yearly_year = st.text_input("在職年年份", value="114", key="add_yearly_year")
                        c14, c15 = st.columns(2)
                        with c14:
                            yearly_amount = st.number_input("在職年金額", min_value=0.0, step=1000.0, format="%.0f", key="add_yearly_amount")
                        with c15:
                            person_note = st.text_input("備註（可空白）", key="add_person_note")
                        submit_person = st.form_submit_button("新增個人資料")
                    if submit_person:
                        if not income_name.strip() or not str(income_project).strip():
                            st.error("請輸入姓名與案名。")
                        else:
                            save_import_records(
                                "個人所得手動",
                                "manual_input_income",
                                [{
                                    "sheet_name": person_company,
                                    "employee_name": income_name.strip(),
                                    "company_name": person_company,
                                    "project_name": str(income_project).strip(),
                                    "roc_year": 114,
                                    "salary": float(income_total),
                                    "bonus": float(income_tax),
                                    "welfare": float(income_health),
                                    "total_income": float(income_net),
                                    "note": append_note_parts([
                                        f"date:{person_date.isoformat()}",
                                        f"所得類型:{income_type}",
                                        f"應扣所得:{float(income_wtax)}",
                                        f"應扣二代:{float(income_whealth)}",
                                        person_note.strip(),
                                    ]),
                                }],
                            )
                            if sync_yearly and yearly_year.strip():
                                annual_amount = float(yearly_amount) if yearly_amount else float(income_total)
                                save_import_records(
                                    "在職年手動",
                                    "manual_input_bonus",
                                    [{
                                        "sheet_name": "在職年統計",
                                        "employee_name": income_name.strip(),
                                        "company_name": person_company,
                                        "project_name": str(income_project).strip(),
                                        "roc_year": 114,
                                        "salary": 0.0,
                                        "bonus": annual_amount,
                                        "welfare": 0.0,
                                        "total_income": annual_amount,
                                        "note": append_note_parts([
                                            f"date:{person_date.isoformat()}",
                                            f"{yearly_year.strip()}:{annual_amount}",
                                            person_note.strip(),
                                        ]),
                                    }],
                                )
                            st.success("已新增個人資料。")
                            st.rerun()

                st.markdown("#### 從原始檔自動帶入六分表")
                default_path = r"d:\data\Downloads\薪資占比1150131.xlsx"
                import_path = st.text_input("來源檔案路徑", value=default_path, key="subsheet_import_path")
                col_i1, col_i2 = st.columns([2, 1])
                with col_i1:
                    replace_imported = st.checkbox("先清除舊的「總表分表匯入」資料再重匯", value=True, key="replace_subsheet_import")
                with col_i2:
                    do_auto_import = st.button("自動帶入六分表", key="auto_import_subsheets")

                if do_auto_import:
                    if not os.path.exists(import_path):
                        st.error("找不到檔案路徑，請確認後再試。")
                    else:
                        if replace_imported:
                            deleted = delete_records_by_source("總表分表匯入")
                            st.info(f"已清除舊匯入資料 {deleted} 筆。")
                        auto_records = parse_total_subsheet_workbook(import_path)
                        if not auto_records:
                            st.warning("沒有偵測到可匯入資料，請確認檔案格式。")
                        else:
                            inserted = save_import_records("總表分表匯入", os.path.basename(import_path), auto_records)
                            st.success(f"自動帶入完成，共 {inserted} 筆。")
                            st.rerun()

                sub_tabs = st.tabs(TOTAL_SUBSHEETS)
                for i, cat in enumerate(TOTAL_SUBSHEETS):
                    with sub_tabs[i]:
                        raw_sub = sub_frames[cat].copy()
                        if raw_sub.empty:
                            st.info("目前沒有資料。")
                        else:
                            editable_sub = raw_sub.iloc[:-1].copy()
                            edited_sub = st.data_editor(
                                editable_sub,
                                use_container_width=True,
                                hide_index=True,
                                key=f"sub_inline_{cat}",
                            )
                            if st.button(f"儲存 {cat} 表格修改", key=f"save_sub_inline_{cat}"):
                                base = editable_sub.copy()
                                now = edited_sub.copy()
                                for col in ["金額"]:
                                    base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0.0)
                                    now[col] = pd.to_numeric(now[col], errors="coerce").fillna(0.0)

                                source_key = f"總表分表調整_{cat}"
                                adjust_records = []
                                for r in range(len(now)):
                                    old_company = str(base.iloc[r]["公司名"])
                                    old_project = str(base.iloc[r]["案場"])
                                    new_company = str(now.iloc[r]["公司名"])
                                    new_project = str(now.iloc[r]["案場"])
                                    old_amount = float(base.iloc[r]["金額"])
                                    new_amount = float(now.iloc[r]["金額"])

                                    if old_company == new_company and old_project == new_project:
                                        delta = new_amount - old_amount
                                        if abs(delta) < 1e-6:
                                            continue
                                        adjust_records.append(
                                            {
                                                "sheet_name": "全案總表分表",
                                                "employee_name": "未指定",
                                                "company_name": new_company,
                                                "project_name": new_project,
                                                "roc_year": None,
                                                "salary": 0.0,
                                                "bonus": 0.0,
                                                "welfare": 0.0,
                                                "total_income": float(delta),
                                                "note": f"category:{cat};inline_sub_adjustment",
                                            }
                                        )
                                    else:
                                        if abs(old_amount) > 1e-6:
                                            adjust_records.append(
                                                {
                                                    "sheet_name": "全案總表分表",
                                                    "employee_name": "未指定",
                                                    "company_name": old_company,
                                                    "project_name": old_project,
                                                    "roc_year": None,
                                                    "salary": 0.0,
                                                    "bonus": 0.0,
                                                    "welfare": 0.0,
                                                    "total_income": -old_amount,
                                                    "note": f"category:{cat};inline_sub_adjustment",
                                                }
                                            )
                                        if abs(new_amount) > 1e-6:
                                            adjust_records.append(
                                                {
                                                    "sheet_name": "全案總表分表",
                                                    "employee_name": "未指定",
                                                    "company_name": new_company,
                                                    "project_name": new_project,
                                                    "roc_year": None,
                                                    "salary": 0.0,
                                                    "bonus": 0.0,
                                                    "welfare": 0.0,
                                                    "total_income": new_amount,
                                                    "note": f"category:{cat};inline_sub_adjustment",
                                                }
                                            )
                                upsert_adjustment_records(source_key, adjust_records)
                                st.success(f"{cat} 已儲存。")
                                st.rerun()

        elif report_view == "在職年統計（薪資獎金）":
            df = df_all[
                df_all["source_type"].isin(["薪資獎金統計", "在職年調整", "在職年手動"])
            ].copy()
            if df.empty:
                st.warning("尚未匯入「薪資獎金統計」資料。")
            else:
                parsed = df["note"].apply(parse_year_amounts)
                years_set = sorted({y for item in parsed for y in item.keys()}, key=lambda x: int(x))
                if not years_set:
                    years_set = ["113", "114", "115"]
                show_years = years_set[:3]

                out = pd.DataFrame({"姓名": df["employee_name"]})
                for y in show_years:
                    out[f"{y}年"] = parsed.apply(lambda d: float(d.get(y, 0.0)))
                out["合計"] = pd.to_numeric(df["total_income"], errors="coerce").fillna(0.0)
                out = out.groupby("姓名", as_index=False)[[f"{y}年" for y in show_years] + ["合計"]].sum()
                editable_out = out.copy()
                edited_out = st.data_editor(editable_out, use_container_width=True, hide_index=True, key="bonus_inline_edit")
                if st.button("儲存在職年統計修改", key="save_bonus_inline"):
                    adjust_records = []
                    for i in range(len(edited_out)):
                        name = str(edited_out.iloc[i]["姓名"])
                        base_vals = {y: float(pd.to_numeric(editable_out.iloc[i][f"{y}年"], errors="coerce")) for y in show_years}
                        new_vals = {y: float(pd.to_numeric(edited_out.iloc[i][f"{y}年"], errors="coerce")) for y in show_years}
                        bsum = float(pd.to_numeric(editable_out.iloc[i]["合計"], errors="coerce"))
                        nsum = float(pd.to_numeric(edited_out.iloc[i]["合計"], errors="coerce"))
                        deltas = {y: new_vals[y] - base_vals[y] for y in show_years}
                        dsum = nsum - bsum
                        if all(abs(v) < 1e-6 for v in deltas.values()) and abs(dsum) < 1e-6:
                            continue
                        adjust_records.append(
                            {
                                "sheet_name": "在職年統計",
                                "employee_name": name,
                                "company_name": None,
                                "project_name": None,
                                "roc_year": 114,
                                "salary": 0.0,
                                "bonus": float(deltas.get(show_years[1], 0.0) if len(show_years) > 1 else list(deltas.values())[0]),
                                "welfare": 0.0,
                                "total_income": dsum,
                                "note": ", ".join([f"{y}:{deltas[y]}" for y in show_years]),
                            }
                        )
                    upsert_adjustment_records("在職年調整", adjust_records)
                    st.success("在職年統計已更新。")
                    st.rerun()
                out = edited_out.copy()
                total_row = pd.DataFrame(
                    [
                        {
                            "姓名": "合計",
                            **{f"{y}年": out[f"{y}年"].sum() for y in show_years},
                            "合計": out["合計"].sum(),
                        }
                    ]
                )
                display_df = pd.concat([out, total_row], ignore_index=True)
                st.download_button(
                    "匯出此報表 Excel",
                    data=to_excel_bytes(
                        display_df,
                        "在職年統計",
                        numeric_cols=[f"{y}年" for y in show_years] + ["合計"],
                        title="在職年統計（薪資獎金）",
                        column_fill_map={
                            **{f"{y}年": "FFF2CC" for y in show_years},
                            "合計": "E2F0D9",
                        },
                    ),
                    file_name="在職年統計_匯出.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_bonus_stat_report",
                )
                st.caption("此表模擬你原檔中的『在職年統計』結構。")

        else:
            df = df_all[
                (df_all["source_type"].isin(["個人所得", "個人所得手動"]))
                | (df_all["source_type"].fillna("").astype(str).str.startswith("個人所得調整_"))
            ].copy()
            if df.empty:
                st.warning("尚未匯入「個人所得」資料。")
            else:
                wh = df["note"].str.extract(r"應扣所得:(?P<w_tax>[-\d.]+),\s*應扣二代:(?P<w_health>[-\d.]+)")
                wh["w_tax"] = pd.to_numeric(wh["w_tax"], errors="coerce").fillna(0.0)
                wh["w_health"] = pd.to_numeric(wh["w_health"], errors="coerce").fillna(0.0)

                base_df = pd.DataFrame(
                    {
                        "分表": df["sheet_name"],
                        "姓名": df["employee_name"],
                        "案場": df["project_name"],
                        "總金額": pd.to_numeric(df["salary"], errors="coerce").fillna(0.0),
                        "所得稅": pd.to_numeric(df["bonus"], errors="coerce").fillna(0.0),
                        "二代健保": pd.to_numeric(df["welfare"], errors="coerce").fillna(0.0),
                        "實領": pd.to_numeric(df["total_income"], errors="coerce").fillna(0.0),
                        "應扣所得": wh["w_tax"],
                        "應扣二代": wh["w_health"],
                    }
                )

                sheet_names = sorted(base_df["分表"].dropna().astype(str).unique().tolist())
                selected_sheet = st.selectbox("網頁顯示分表", sheet_names, index=0)
                out = base_df[base_df["分表"] == selected_sheet].copy()
                out = out.groupby(["姓名", "案場"], as_index=False)[["總金額", "所得稅", "二代健保", "實領", "應扣所得", "應扣二代"]].sum()
                editable_out = out.copy()
                edited_out = st.data_editor(editable_out, use_container_width=True, hide_index=True, key="income_inline_edit")
                if st.button("儲存個人所得修改", key="save_income_inline"):
                    source_key = f"個人所得調整_{selected_sheet}"
                    adjust_records = []
                    for i in range(len(edited_out)):
                        name = str(edited_out.iloc[i]["姓名"])
                        project = str(edited_out.iloc[i]["案場"])
                        b = editable_out.iloc[i]
                        n = edited_out.iloc[i]
                        d_total = float(pd.to_numeric(n["總金額"], errors="coerce") - pd.to_numeric(b["總金額"], errors="coerce"))
                        d_tax = float(pd.to_numeric(n["所得稅"], errors="coerce") - pd.to_numeric(b["所得稅"], errors="coerce"))
                        d_health = float(pd.to_numeric(n["二代健保"], errors="coerce") - pd.to_numeric(b["二代健保"], errors="coerce"))
                        d_net = float(pd.to_numeric(n["實領"], errors="coerce") - pd.to_numeric(b["實領"], errors="coerce"))
                        d_wtax = float(pd.to_numeric(n["應扣所得"], errors="coerce") - pd.to_numeric(b["應扣所得"], errors="coerce"))
                        d_whealth = float(pd.to_numeric(n["應扣二代"], errors="coerce") - pd.to_numeric(b["應扣二代"], errors="coerce"))
                        if all(abs(v) < 1e-6 for v in [d_total, d_tax, d_health, d_net, d_wtax, d_whealth]):
                            continue
                        adjust_records.append(
                            {
                                "sheet_name": selected_sheet,
                                "employee_name": name,
                                "company_name": selected_sheet,
                                "project_name": project,
                                "roc_year": 114,
                                "salary": d_total,
                                "bonus": d_tax,
                                "welfare": d_health,
                                "total_income": d_net,
                                "note": f"應扣所得:{d_wtax}, 應扣二代:{d_whealth}",
                            }
                        )
                    upsert_adjustment_records(source_key, adjust_records)
                    st.success("個人所得已更新。")
                    st.rerun()
                total_row = pd.DataFrame(
                    [
                        {
                            "姓名": "合計",
                            "案場": "",
                            "總金額": out["總金額"].sum(),
                            "所得稅": out["所得稅"].sum(),
                            "二代健保": out["二代健保"].sum(),
                            "實領": out["實領"].sum(),
                            "應扣所得": out["應扣所得"].sum(),
                            "應扣二代": out["應扣二代"].sum(),
                        }
                    ]
                )
                display_df = pd.concat([out, total_row], ignore_index=True)

                export_sheets: dict[str, pd.DataFrame] = {}
                for sheet_name in sheet_names:
                    sheet_df = base_df[base_df["分表"] == sheet_name].copy()
                    sheet_df = sheet_df.groupby(["姓名", "案場"], as_index=False)[
                        ["總金額", "所得稅", "二代健保", "實領", "應扣所得", "應扣二代"]
                    ].sum()
                    sheet_total = pd.DataFrame(
                        [
                            {
                                "姓名": "合計",
                                "案場": "",
                                "總金額": sheet_df["總金額"].sum(),
                                "所得稅": sheet_df["所得稅"].sum(),
                                "二代健保": sheet_df["二代健保"].sum(),
                                "實領": sheet_df["實領"].sum(),
                                "應扣所得": sheet_df["應扣所得"].sum(),
                                "應扣二代": sheet_df["應扣二代"].sum(),
                            }
                        ]
                    )
                    export_sheets[sheet_name] = pd.concat([sheet_df, sheet_total], ignore_index=True)

                st.download_button(
                    "匯出全部分表 Excel（多工作表）",
                    data=to_excel_multi_sheets(
                        sheet_data=export_sheets,
                        numeric_cols=["總金額", "所得稅", "二代健保", "實領", "應扣所得", "應扣二代"],
                        title_prefix="個人所得明細",
                        column_fill_map={
                            "總金額": "FFF2CC",
                            "所得稅": "FFF2CC",
                            "二代健保": "FFF2CC",
                            "實領": "E2F0D9",
                            "應扣所得": "EDEDED",
                            "應扣二代": "EDEDED",
                        },
                    ),
                    file_name="個人所得_多分表匯出.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_personal_income_report",
                )
                st.caption("網頁可切換分表，匯出時會產生同一檔案的多工作表（底下可切換分頁）。")

with tab_query:
    st.subheader("已匯入資料查詢")
    q1, q2, q3 = st.columns([2, 1, 1])
    with q1:
        keyword = st.text_input("關鍵字（姓名/公司/案場）", placeholder="輸入關鍵字")
    with q2:
        source_type = st.selectbox("來源", ["全部", "薪資占比", "薪資獎金統計", "個人所得"])
    with q3:
        year_input = st.text_input("年度(民國)", placeholder="例如 114")
    roc_year = int(year_input) if year_input.strip().isdigit() else None

    rows = list_payroll_records(keyword=keyword, source_type=source_type, roc_year=roc_year)
    if rows:
        df = pd.DataFrame([dict(r) for r in rows])
        edit_cols = [
            "id",
            "source_type",
            "sheet_name",
            "employee_name",
            "company_name",
            "project_name",
            "roc_year",
            "salary",
            "bonus",
            "welfare",
            "total_income",
            "note",
        ]
        editable_df = df[edit_cols].copy()
        edited = st.data_editor(editable_df, use_container_width=True, hide_index=True, key="query_inline_editor")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("儲存目前編輯", key="save_inline_edit"):
                for _, row in edited.iterrows():
                    rid = int(row["id"])
                    update_payroll_record(
                        rid,
                        {
                            "sheet_name": row["sheet_name"],
                            "employee_name": row["employee_name"],
                            "company_name": row["company_name"],
                            "project_name": row["project_name"],
                            "roc_year": None if pd.isna(row["roc_year"]) else int(row["roc_year"]),
                            "salary": float(row["salary"] or 0),
                            "bonus": float(row["bonus"] or 0),
                            "welfare": float(row["welfare"] or 0),
                            "total_income": float(row["total_income"] or 0),
                            "note": "" if pd.isna(row["note"]) else str(row["note"]),
                        },
                    )
                st.success("已儲存修改。")
                st.rerun()
        with c2:
            delete_ids_text = st.text_input("刪除ID（逗號分隔）", placeholder="例如 12,15", key="query_delete_ids")
            if st.button("刪除指定ID", key="delete_inline_ids"):
                ids = [i.strip() for i in delete_ids_text.split(",") if i.strip()]
                deleted = delete_payroll_records([int(i) for i in ids if i.isdigit()])
                st.success(f"已刪除 {deleted} 筆。")
                st.rerun()
    else:
        st.info("查無資料。")

with tab_batches:
    st.subheader("匯入批次紀錄")
    batches = list_batches()
    if batches:
        st.dataframe(pd.DataFrame([dict(r) for r in batches]), use_container_width=True, hide_index=True)
    else:
        st.info("尚無匯入紀錄。")

    st.markdown("### 資料庫備份與還原")
    db_file = Path(DB_PATH)
    if db_file.exists():
        with db_file.open("rb") as f:
            st.download_button(
                "下載資料庫備份（.db）",
                data=f.read(),
                file_name=f"financial_reports_backup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.db",
                mime="application/octet-stream",
                key="download_db_backup",
            )
    else:
        st.info("目前尚未建立資料庫檔案，先新增或匯入資料後再備份。")

    restore_file = st.file_uploader("上傳備份檔還原（.db）", type=["db"], key="restore_db_uploader")
    if restore_file is not None:
        st.warning("還原會覆蓋目前資料庫，請先下載備份。")
        if st.button("確認還原資料庫", key="confirm_restore_db"):
            try:
                db_file.parent.mkdir(parents=True, exist_ok=True)
                with db_file.open("wb") as f:
                    f.write(restore_file.getvalue())
                st.success("資料庫已還原，請重新整理頁面。")
                st.rerun()
            except Exception as exc:
                st.error(f"還原失敗：{exc}")
