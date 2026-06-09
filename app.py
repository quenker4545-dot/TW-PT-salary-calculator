from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import BytesIO
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# =========================================================
# 台灣 PT / 短期勞工薪資試算器
# ---------------------------------------------------------
# 注意：
# 1. 勞健保級距、費率會隨年度異動，請在側欄確認。
# 2. 本工具是薪資估算器，不是法律意見。
# 3. 真正薪資仍以雇主申報之投保薪資、薪資單、勞保局/健保署資料為準。
# =========================================================

st.set_page_config(
    page_title="台灣 PT 實拿薪資計算機",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- UI CSS: mobile-friendly ----------
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
    div[data-testid="stMetricValue"] { font-size: 1.55rem; }
    div[data-testid="stMetricLabel"] { font-size: 0.95rem; }
    .small-note { color: #666; font-size: 0.9rem; }
    .risk-box {
        border: 1px solid rgba(250, 150, 50, 0.35);
        background: rgba(250, 150, 50, 0.08);
        border-radius: 10px;
        padding: 0.8rem 1rem;
        margin: 0.5rem 0;
    }
    .ok-box {
        border: 1px solid rgba(60, 160, 90, 0.35);
        background: rgba(60, 160, 90, 0.08);
        border-radius: 10px;
        padding: 0.8rem 1rem;
        margin: 0.5rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------- Defaults: user-editable ----------
# 115年/2026常見低級距。若官方調整，請直接改這裡或在 UI 手動指定投保薪資。
LABOR_INSURED_GRADES = [
    29500, 30300, 31800, 33300, 34800, 36300,
    38200, 40100, 42000, 43900, 45800,
]

# 健保常見級距，PT 通常落在前面；保留高一點的級距避免超過 45,800 時失真太多。
NHI_INSURED_GRADES = [
    29500, 30300, 31800, 33300, 34800, 36300, 38200, 40100, 42000, 43900, 45800,
    48200, 50600, 53000, 55400, 57800, 60800, 63800, 66800, 69800, 72800,
    76500, 80200, 83900, 87600, 92100, 96600, 101100, 105600, 110100,
    115500, 120900, 126300, 131700, 137100, 142500, 147900, 150000,
    156400, 162800, 169200, 175600, 182000, 189500, 197000, 204500,
    212000, 219500, 228200, 236900, 245600, 254300, 263000, 276000,
    289000, 303000, 313000,
]

DAY_TYPES = [
    "平日",
    "休息日",
    "國定假日/雙薪日",
    "例假日/特殊出勤",
    "不計加班(純時薪)",
]


# ---------- Utility ----------
def money(x: float | int | Decimal) -> int:
    """Round to nearest NT dollar using ROUND_HALF_UP."""
    return int(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def fmt(n: float | int) -> str:
    return f"NT$ {money(n):,}"


def ceil_to_grade(amount: float, grades: List[int]) -> int:
    if amount <= 0:
        return grades[0]
    for g in grades:
        if amount <= g:
            return g
    return grades[-1]


@dataclass
class DayPay:
    regular_hours: float
    ot_134_hours: float
    ot_167_hours: float
    ot_267_hours: float
    regular_pay: float
    overtime_pay: float
    total_pay: float
    warning: str = ""


def calc_day_pay(hours: float, hourly_rate: float, day_type: str) -> DayPay:
    """Calculate one day's wage.

    Multipliers are represented as total hourly pay multipliers:
    - 1.3333 = hourly wage + 1/3
    - 1.6667 = hourly wage + 2/3
    - 2.6667 = hourly wage + 1 + 2/3
    """
    h = max(float(hours), 0.0)
    r = float(hourly_rate)
    warn = ""

    if h > 12 and day_type in ["平日", "休息日", "國定假日/雙薪日"]:
        warn = "單日工時超過 12 小時；請確認是否合法與是否另有特殊約定。"

    if day_type == "平日":
        regular_h = min(h, 8)
        ot134 = max(min(h - 8, 2), 0)
        ot167 = max(h - 10, 0)
        regular_pay = regular_h * r
        overtime_pay = ot134 * r * (4 / 3) + ot167 * r * (5 / 3)
        return DayPay(regular_h, ot134, ot167, 0, regular_pay, overtime_pay, regular_pay + overtime_pay, warn)

    if day_type == "休息日":
        # 休息日出勤全數以休息日加班費率計算。
        ot134 = min(h, 2)
        ot167 = max(min(h - 2, 6), 0)
        ot267 = max(h - 8, 0)
        overtime_pay = ot134 * r * (4 / 3) + ot167 * r * (5 / 3) + ot267 * r * (8 / 3)
        return DayPay(0, ot134, ot167, ot267, 0, overtime_pay, overtime_pay, warn)

    if day_type == "國定假日/雙薪日":
        # 常見口語「雙薪」：8 小時內按 2 倍；超過 8 小時用平日延長工時倍率估算。
        regular_h = min(h, 8)
        ot134 = max(min(h - 8, 2), 0)
        ot167 = max(h - 10, 0)
        regular_pay = regular_h * r * 2
        overtime_pay = ot134 * r * (4 / 3) + ot167 * r * (5 / 3)
        return DayPay(regular_h, ot134, ot167, 0, regular_pay, overtime_pay, regular_pay + overtime_pay, warn)

    if day_type == "例假日/特殊出勤":
        # 例假日原則不得任意出勤；此處只做保守雙倍估算並提示風險。
        warn = "例假日原則上不能任意出勤；此列僅用 2 倍薪資估算，請確認實際法定情境。"
        total = h * r * 2
        return DayPay(0, 0, 0, 0, 0, total, total, warn)

    # 不計加班：用於對照某些活動工讀公司只用「時薪 x 工時」給薪。
    regular_pay = h * r
    return DayPay(h, 0, 0, 0, regular_pay, 0, regular_pay, warn)


def calc_payroll(detail_df: pd.DataFrame, hourly_rate: float) -> Tuple[pd.DataFrame, Dict[str, float], List[str]]:
    rows: List[Dict[str, Any]] = []
    warnings: List[str] = []

    for _, row in detail_df.iterrows():
        day_no = int(row["天數"])
        hours = float(row["工時"])
        day_type = str(row["日別"])
        note = str(row.get("備註", ""))

        dp = calc_day_pay(hours, hourly_rate, day_type)
        if dp.warning:
            warnings.append(f"第 {day_no} 天：{dp.warning}")

        raw_pay = hours * hourly_rate
        rows.append({
            "天數": day_no,
            "日別": day_type,
            "工時": hours,
            "正常/雙薪工時": dp.regular_hours,
            "1.34倍時數": dp.ot_134_hours,
            "1.67倍時數": dp.ot_167_hours,
            "2.67倍時數": dp.ot_267_hours,
            "正常/雙薪工資": money(dp.regular_pay),
            "加班/特殊工資": money(dp.overtime_pay),
            "依法/約定應領": money(dp.total_pay),
            "純時薪算法": money(raw_pay),
            "差額": money(dp.total_pay - raw_pay),
            "備註": note if note != "nan" else "",
        })

    out = pd.DataFrame(rows)
    summary = {
        "total_hours": float(out["工時"].sum()) if not out.empty else 0.0,
        "gross_pay": float(out["依法/約定應領"].sum()) if not out.empty else 0.0,
        "raw_pay": float(out["純時薪算法"].sum()) if not out.empty else 0.0,
        "difference": float(out["差額"].sum()) if not out.empty else 0.0,
        "work_days": float(len(out)),
    }
    return out, summary, warnings


def calc_labor_fee(
    gross_pay: float,
    insured_salary: int,
    insured_days: int,
    labor_total_rate: float,
    employee_share: float,
    prorate: bool,
) -> int:
    if insured_salary <= 0:
        return 0
    day_ratio = max(min(insured_days, 30), 0) / 30 if prorate else 1
    return money(insured_salary * labor_total_rate * employee_share * day_ratio)


def calc_nhi_fee(
    insured_salary: int,
    nhi_rate: float,
    employee_share: float,
    dependents: int,
    prorate: bool,
    insured_days: int,
) -> int:
    if insured_salary <= 0:
        return 0
    # 本人 + 眷屬，眷屬最多通常以 3 口計。
    persons = 1 + max(min(int(dependents), 3), 0)
    day_ratio = max(min(insured_days, 30), 0) / 30 if prorate else 1
    return money(insured_salary * nhi_rate * employee_share * persons * day_ratio)


def build_excel(detail: pd.DataFrame, summary_rows: List[Tuple[str, Any]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "薪資明細"

    title_fill = PatternFill("solid", fgColor="1F4E78")
    title_font = Font(color="FFFFFF", bold=True)
    header_fill = PatternFill("solid", fgColor="D9EAF7")
    bold_font = Font(bold=True)

    ws["A1"] = "台灣 PT 實拿薪資計算結果"
    ws["A1"].font = Font(bold=True, size=16)
    ws["A2"] = f"產生日期：{date.today().isoformat()}"

    start_row = 4
    for col_idx, col_name in enumerate(detail.columns, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=col_name)
        cell.fill = header_fill
        cell.font = bold_font
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in enumerate(detail.itertuples(index=False), start=start_row + 1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    summary_start = start_row + len(detail) + 3
    ws.cell(row=summary_start, column=1, value="摘要").fill = title_fill
    ws.cell(row=summary_start, column=1).font = title_font

    for idx, (k, v) in enumerate(summary_rows, start=summary_start + 1):
        ws.cell(row=idx, column=1, value=k).font = bold_font
        ws.cell(row=idx, column=2, value=v)

    for col in range(1, len(detail.columns) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 18

    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()


# ---------- App ----------
st.title("台灣 PT／短期勞工實拿薪資計算機")
st.caption("支援平日加班、休息日、國定假日/雙薪日、勞保、健保、所得稅預扣、Excel 匯出。")

with st.expander("使用前先看：這個工具的邊界", expanded=False):
    st.markdown(
        """
        - 這是**估算器**，實際扣款仍取決於公司申報的投保薪資級距、加退保日期、薪資單算法。
        - 勞退 6% 是雇主提繳，**不應從你的薪水扣**，所以本工具只顯示參考值，不列入扣款。
        - 若你本來已有正職健保，兼職通常不一定會再扣健保；請用側欄把健保關掉或改成手動估算。
        """
    )

# ---------- Sidebar settings ----------
with st.sidebar:
    st.header("基本設定")
    hourly_rate = st.number_input("時薪（NTD）", min_value=0.0, value=250.0, step=10.0)
    work_days = st.number_input("工作天數", min_value=1, max_value=31, value=5, step=1)

    st.divider()
    st.header("保險與扣款")
    insurance_mode = st.selectbox(
        "投保模式",
        ["短期臨時工/活動工讀", "長期兼職/當月薪資", "自訂"],
        index=0,
    )

    has_labor = st.checkbox("扣勞保 + 就保自付額", value=True)
    has_nhi = st.checkbox("扣健保自付額", value=False)
    withhold_tax = st.checkbox("所得稅預扣 5%", value=False)

    if insurance_mode == "短期臨時工/活動工讀":
        default_prorate_labor = True
        default_prorate_nhi = True
        default_insured_days = int(work_days)
    elif insurance_mode == "長期兼職/當月薪資":
        default_prorate_labor = False
        default_prorate_nhi = False
        default_insured_days = 30
    else:
        default_prorate_labor = True
        default_prorate_nhi = False
        default_insured_days = int(work_days)

    insured_days = st.number_input("加保天數（估算用）", min_value=1, max_value=30, value=default_insured_days, step=1)
    prorate_labor = st.checkbox("勞保按加保天數比例估算", value=default_prorate_labor)
    prorate_nhi = st.checkbox("健保按加保天數比例估算", value=default_prorate_nhi)

    dependents = st.number_input("健保眷屬人數（不含本人，最多估 3 人）", min_value=0, max_value=3, value=0, step=1)

    st.divider()
    st.header("費率/級距")
    st.caption("若官方費率調整，直接改這裡即可。")
    labor_total_rate = st.number_input("勞保+就保合計費率", min_value=0.0, max_value=0.3, value=0.125, step=0.001, format="%.3f")
    labor_employee_share = st.number_input("勞保勞工負擔比例", min_value=0.0, max_value=1.0, value=0.20, step=0.01, format="%.2f")
    nhi_rate = st.number_input("健保費率", min_value=0.0, max_value=0.2, value=0.0517, step=0.0001, format="%.4f")
    nhi_employee_share = st.number_input("健保本人負擔比例", min_value=0.0, max_value=1.0, value=0.30, step=0.01, format="%.2f")

    st.divider()
    manual_insured_salary = st.checkbox("手動指定投保薪資", value=False)
    if manual_insured_salary:
        labor_insured_salary_manual = st.number_input("勞保月投保薪資", min_value=0, value=29500, step=100)
        nhi_insured_salary_manual = st.number_input("健保月投保金額", min_value=0, value=29500, step=100)

# ---------- Day editor ----------
def default_schedule(n: int) -> pd.DataFrame:
    return pd.DataFrame({
        "天數": list(range(1, n + 1)),
        "日別": ["平日"] * n,
        "工時": [8.0] * n,
        "備註": [""] * n,
    })

st.subheader("每日工時與日別")
st.markdown("把每天的工時和日別填進去；如果公司說某一天是雙倍薪資，日別選 **國定假日/雙薪日**。")

input_df = default_schedule(int(work_days))
edited_df = st.data_editor(
    input_df,
    hide_index=True,
    use_container_width=True,
    column_config={
        "天數": st.column_config.NumberColumn("天數", disabled=True, width="small"),
        "日別": st.column_config.SelectboxColumn("日別", options=DAY_TYPES, required=True, width="medium"),
        "工時": st.column_config.NumberColumn("工時", min_value=0.0, max_value=24.0, step=0.5, required=True, width="small"),
        "備註": st.column_config.TextColumn("備註", width="medium"),
    },
    key="schedule_editor",
)

# ---------- Compute ----------
detail, pay_summary, warnings = calc_payroll(edited_df, hourly_rate)
gross_pay = pay_summary["gross_pay"]
raw_pay = pay_summary["raw_pay"]
diff_pay = pay_summary["difference"]
total_hours = pay_summary["total_hours"]

# Insurance salary estimation: for short-term daily work, estimate monthly equivalent using average daily wage x 30.
if insurance_mode == "短期臨時工/活動工讀":
    avg_daily_pay = gross_pay / max(float(work_days), 1)
    monthly_equivalent = avg_daily_pay * 30
else:
    monthly_equivalent = gross_pay

if manual_insured_salary:
    labor_insured_salary = int(labor_insured_salary_manual)
    nhi_insured_salary = int(nhi_insured_salary_manual)
else:
    labor_insured_salary = ceil_to_grade(monthly_equivalent, LABOR_INSURED_GRADES)
    nhi_insured_salary = ceil_to_grade(monthly_equivalent, NHI_INSURED_GRADES)

labor_fee = calc_labor_fee(
    gross_pay=gross_pay,
    insured_salary=labor_insured_salary,
    insured_days=int(insured_days),
    labor_total_rate=float(labor_total_rate),
    employee_share=float(labor_employee_share),
    prorate=bool(prorate_labor),
) if has_labor else 0

nhi_fee = calc_nhi_fee(
    insured_salary=nhi_insured_salary,
    nhi_rate=float(nhi_rate),
    employee_share=float(nhi_employee_share),
    dependents=int(dependents),
    prorate=bool(prorate_nhi),
    insured_days=int(insured_days),
) if has_nhi else 0

tax_fee = money(gross_pay * 0.05) if withhold_tax else 0
labor_pension_reference = money(gross_pay * 0.06)
total_deductions = labor_fee + nhi_fee + tax_fee
net_pay = money(gross_pay - total_deductions)

# ---------- Metrics ----------
st.subheader("計算結果")
metric_cols = st.columns(4)
metric_cols[0].metric("應領薪資", fmt(gross_pay))
metric_cols[1].metric("扣款合計", fmt(total_deductions))
metric_cols[2].metric("預估實拿", fmt(net_pay))
metric_cols[3].metric("相對純時薪差額", fmt(diff_pay))

if diff_pay > 0:
    st.markdown(
        f"<div class='ok-box'>若公司只用 <b>時薪 × 總工時</b> 計算，會是 {fmt(raw_pay)}；本工具估算應領 {fmt(gross_pay)}，差額約 <b>{fmt(diff_pay)}</b>。</div>",
        unsafe_allow_html=True,
    )
elif diff_pay < 0:
    st.markdown(
        f"<div class='risk-box'>目前設定下，依法/約定應領低於純時薪算法，請確認日別設定是否正確。差額：{fmt(diff_pay)}</div>",
        unsafe_allow_html=True,
    )

if warnings:
    for w in warnings:
        st.warning(w)

# ---------- Tables ----------
tab1, tab2, tab3 = st.tabs(["薪資明細", "扣款拆解", "圖表/匯出"])

with tab1:
    st.dataframe(detail, use_container_width=True, hide_index=True)

with tab2:
    deduction_df = pd.DataFrame([
        {"項目": "勞保+就保自付額", "是否啟用": "是" if has_labor else "否", "投保薪資/金額": labor_insured_salary if has_labor else 0, "扣款": labor_fee, "備註": "勞工自付估算；職災保險通常雇主負擔"},
        {"項目": "健保自付額", "是否啟用": "是" if has_nhi else "否", "投保薪資/金額": nhi_insured_salary if has_nhi else 0, "扣款": nhi_fee, "備註": f"本人+眷屬 {1 + int(dependents)} 人估算"},
        {"項目": "所得稅預扣", "是否啟用": "是" if withhold_tax else "否", "投保薪資/金額": "-", "扣款": tax_fee, "備註": "常見預扣 5%，是否預扣依實際給付型態"},
        {"項目": "勞退6%參考", "是否啟用": "雇主提繳", "投保薪資/金額": "-", "扣款": 0, "備註": f"雇主約應提繳 {fmt(labor_pension_reference)}，不應從薪資扣"},
    ])
    st.dataframe(deduction_df, use_container_width=True, hide_index=True)

    st.markdown(
        f"""
        **保險估算基礎**

        - 投保模式：`{insurance_mode}`
        - 估算月薪基礎：`{fmt(monthly_equivalent)}`
        - 勞保月投保薪資：`{labor_insured_salary:,}`
        - 健保月投保金額：`{nhi_insured_salary:,}`
        - 加保天數：`{int(insured_days)}`
        """
    )

with tab3:
    chart_df = pd.DataFrame({
        "項目": ["應領薪資", "勞保", "健保", "所得稅", "實拿"],
        "金額": [money(gross_pay), labor_fee, nhi_fee, tax_fee, net_pay],
    }).set_index("項目")
    st.bar_chart(chart_df)

    summary_rows = [
        ("時薪", money(hourly_rate)),
        ("總工時", total_hours),
        ("純時薪算法", money(raw_pay)),
        ("應領薪資", money(gross_pay)),
        ("勞保扣款", labor_fee),
        ("健保扣款", nhi_fee),
        ("所得稅預扣", tax_fee),
        ("扣款合計", total_deductions),
        ("預估實拿", net_pay),
        ("勞退6%參考_雇主提繳非扣款", labor_pension_reference),
    ]

    excel_bytes = build_excel(detail, summary_rows)
    st.download_button(
        "下載 Excel 薪資明細",
        data=excel_bytes,
        file_name=f"pt_salary_result_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.divider()
st.caption("本工具僅供估算。若要用於正式請款或勞資爭議，請以薪資單、投保資料、主管機關公告與個案契約為準。")
