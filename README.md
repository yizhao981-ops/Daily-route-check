import io
import datetime
import pytz
import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

TZ = "America/New_York"


def detect_col(df: pd.DataFrame, key: str):
    key = key.upper()
    for c in df.columns:
        if key in str(c).upper():
            return c
    return None


def normalize_status(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip().str.upper()


def style_header(ws, header_row: int = 1):
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[header_row]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = f"A{header_row + 1}"
    ws.auto_filter.ref = ws.dimensions


def autosize(ws, cap=45, sample_rows=500):
    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_len = 10
        for cell in col_cells[:sample_rows]:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, cap)


def apply_route_colors(ws, colnames, header_row: int = 1):
    yellow = PatternFill("solid", fgColor="FFF2CC")
    red = PatternFill("solid", fgColor="F8CBAD")
    purple = PatternFill("solid", fgColor="E4DFEC")

    min_idx = colnames.index("MinutesSinceLast") + 1
    flag_idx = colnames.index("StatusFlag") + 1
    rate_idx = colnames.index("OFD(CompletionRate)") + 1

    for r in range(header_row + 1, ws.max_row + 1):
        ws.cell(r, rate_idx).number_format = "0.00%"
        minutes = ws.cell(r, min_idx).value
        flag = ws.cell(r, flag_idx).value

        fill = None
        if flag == "NO_DELIVERED":
            fill = purple
        elif minutes is not None:
            try:
                m = float(minutes)
                if m > 60:
                    fill = red
                elif m > 30:
                    fill = yellow
            except Exception:
                pass

        if fill:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = fill


def write_check_sheet(wb, sheet_name, route_df, route_cols, now_et, hour_threshold, rate_threshold, fill_color):
    ws = wb.create_sheet(sheet_name)
    rule_applied = now_et.hour >= hour_threshold
    flagged = route_df[route_df["OFD(CompletionRate)"] < rate_threshold].copy() if rule_applied else route_df.iloc[0:0].copy()

    ws["A1"] = "RunTime (ET)"
    ws["B1"] = now_et.strftime("%Y-%m-%d %H:%M:%S")
    ws["A2"] = "Rule"
    ws["B2"] = f"At/after {hour_threshold}:00 ET: OFD(CompletionRate) < {int(rate_threshold * 100)}%"
    ws["A3"] = "RuleApplied"
    ws["B3"] = "YES" if rule_applied else f"NO (before {hour_threshold}:00 ET)"

    ws.append([])
    ws.append(route_cols)

    for row in flagged[route_cols].itertuples(index=False):
        ws.append(list(row))

    header_row = 5
    style_header(ws, header_row=header_row)

    rate_idx = route_cols.index("OFD(CompletionRate)") + 1
    for r in range(header_row + 1, ws.max_row + 1):
        ws.cell(r, rate_idx).number_format = "0.00%"

    fill = PatternFill("solid", fgColor=fill_color)
    for r in range(header_row + 1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).fill = fill

    autosize(ws)


def build_excel_bytes(raw_df: pd.DataFrame) -> tuple[bytes, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    df = raw_df.copy()

    if len(df.columns) < 12:
        raise ValueError("Input file must have at least 12 columns because B, J, and L are required.")

    route_col = df.columns[1]   # B
    status_col = df.columns[9]  # J
    time_col = df.columns[11]   # L

    flee_col = detect_col(df, "FLEE")
    driver_col = detect_col(df, "DRIVER")

    df[time_col] = pd.to_datetime(df[time_col], errors="coerce")
    df["StatusU"] = normalize_status(df[status_col])

    tz = pytz.timezone(TZ)
    now_et = datetime.datetime.now(tz)
    now_et_naive = now_et.replace(tzinfo=None)

    rows = []

    for route, g in df.groupby(route_col, dropna=True):
        total = int(len(g))
        success = int((g["StatusU"] == "DELIVERED").sum())
        failed = int(g["StatusU"].str.contains("FAIL", na=False).sum())
        new_count = int((g["StatusU"] == "NEW").sum())
        remaining_ofd = int((g["StatusU"] == "OUT_FOR_DELIVERY").sum())

        ofd_total = int(total - new_count)
        ofd_completion_rate = (success / ofd_total) if ofd_total > 0 else 0.0

        flee = g[flee_col].dropna().iloc[0] if flee_col and g[flee_col].notna().any() else None
        driver = g[driver_col].dropna().iloc[0] if driver_col and g[driver_col].notna().any() else None

        delivered_rows = g[(g["StatusU"] == "DELIVERED") & g[time_col].notna()]

        if delivered_rows.empty:
            first_del = None
            last_del = None
            minutes_since_last = None
            hours_since_first = None
            per_hour = None
            status_flag = "NO_DELIVERED"
            bucket = "NO_DELIVERED"
        else:
            first_del = delivered_rows[time_col].min()
            last_del = delivered_rows[time_col].max()
            minutes_since_last = (now_et_naive - last_del).total_seconds() / 60
            hours_since_first = (now_et_naive - first_del).total_seconds() / 3600
            per_hour = (success / hours_since_first) if hours_since_first and hours_since_first > 0 else None
            status_flag = "HAS_DELIVERED"
            if minutes_since_last > 60:
                bucket = "RED"
            elif minutes_since_last > 30:
                bucket = "YELLOW"
            else:
                bucket = "OK"

        rows.append({
            "Route": route,
            "DriverName": driver,
            "FleeName": flee,
            "Total": total,
            "OFD Total": ofd_total,
            "Success(Delivered)": success,
            "Failed(*FAIL*)": failed,
            "New": new_count,
            "Remaining(OFD)": remaining_ofd,
            "OFD(CompletionRate)": ofd_completion_rate,
            "1stDeliveryTime": first_del,
            "HoursSinceFirstDelivery": round(hours_since_first, 2) if hours_since_first is not None else None,
            "DeliveriesPerHour": round(per_hour, 2) if per_hour is not None else None,
            "LatestDeliveredTime": last_del,
            "MinutesSinceLast": round(minutes_since_last, 1) if minutes_since_last is not None else None,
            "StatusFlag": status_flag,
            "AlertBucket": bucket
        })

    route_df = pd.DataFrame(rows)

    route_df["_sort"] = route_df["MinutesSinceLast"].fillna(10**9)
    route_df.sort_values(["StatusFlag", "_sort"], ascending=[True, False], inplace=True)
    route_df.drop(columns="_sort", inplace=True)

    sum_df = route_df.copy()
    sum_df["FleeName"] = sum_df["FleeName"].fillna("UNKNOWN")

    summary_df = sum_df.groupby("FleeName").agg(
        Routes=("Route", "nunique"),
        TotalPkgs=("Total", "sum"),
        OFDTotal=("OFD Total", "sum"),
        Delivered=("Success(Delivered)", "sum"),
        Failed=("Failed(*FAIL*)", "sum"),
        NewPkgs=("New", "sum"),
        RemainingOFD=("Remaining(OFD)", "sum"),
        NoDeliveredRoutes=("StatusFlag", lambda s: int((s == "NO_DELIVERED").sum())),
        RedRoutes=("AlertBucket", lambda s: int((s == "RED").sum())),
        YellowRoutes=("AlertBucket", lambda s: int((s == "YELLOW").sum())),
        AvgDeliveriesPerHour=("DeliveriesPerHour", "mean"),
    ).reset_index()

    summary_df["OFD(CompletionRate)"] = (
        summary_df["Delivered"] / summary_df["OFDTotal"].replace({0: pd.NA})
    ).fillna(0.0)
    summary_df["AvgDeliveriesPerHour"] = summary_df["AvgDeliveriesPerHour"].round(2)

    summary_df["_urgency"] = (
        summary_df["NoDeliveredRoutes"] * 1000000
        + summary_df["RedRoutes"] * 1000
        + summary_df["YellowRoutes"]
    )
    summary_df.sort_values(["_urgency", "OFD(CompletionRate)"], ascending=[False, True], inplace=True)
    summary_df.drop(columns="_urgency", inplace=True)

    exc_df = route_df[
        (route_df["StatusFlag"] == "NO_DELIVERED") |
        ((route_df["MinutesSinceLast"].fillna(0) > 120) & (route_df["Remaining(OFD)"] > 0)) |
        ((route_df["DeliveriesPerHour"].fillna(999) < 10) & (route_df["Remaining(OFD)"] > 0))
    ].copy()

    wb = Workbook()

    route_cols = [
        "Route", "DriverName", "FleeName", "Total", "OFD Total",
        "Success(Delivered)", "Failed(*FAIL*)", "New", "Remaining(OFD)",
        "OFD(CompletionRate)", "1stDeliveryTime", "HoursSinceFirstDelivery",
        "DeliveriesPerHour", "LatestDeliveredTime", "MinutesSinceLast",
        "StatusFlag", "AlertBucket"
    ]

    ws1 = wb.active
    ws1.title = "RouteMonitor"
    ws1.append(route_cols)
    for row in route_df[route_cols].itertuples(index=False):
        ws1.append(list(row))
    style_header(ws1)
    apply_route_colors(ws1, route_cols)
    autosize(ws1)

    ws2 = wb.create_sheet("Summary")
    summary_cols = [
        "FleeName", "Routes", "TotalPkgs", "OFDTotal", "Delivered",
        "Failed", "NewPkgs", "RemainingOFD", "OFD(CompletionRate)",
        "NoDeliveredRoutes", "RedRoutes", "YellowRoutes", "AvgDeliveriesPerHour"
    ]
    ws2.append(summary_cols)
    for row in summary_df[summary_cols].itertuples(index=False):
        ws2.append(list(row))
    style_header(ws2)
    summary_rate_idx = summary_cols.index("OFD(CompletionRate)") + 1
    for r in range(2, ws2.max_row + 1):
        ws2.cell(r, summary_rate_idx).number_format = "0.00%"
    autosize(ws2)

    ws3 = wb.create_sheet("Exceptions")
    ws3.append(route_cols)
    for row in exc_df[route_cols].itertuples(index=False):
        ws3.append(list(row))
    style_header(ws3)
    apply_route_colors(ws3, route_cols)
    autosize(ws3)

    write_check_sheet(wb, "3pm check", route_df, route_cols, now_et, 15, 0.50, "FCE4D6")
    write_check_sheet(wb, "6pm check", route_df, route_cols, now_et, 18, 0.80, "F4B084")

    ws6 = wb.create_sheet("Meta")
    ws6["A1"] = "Now (ET) used for calculation"
    ws6["B1"] = now_et.strftime("%Y-%m-%d %H:%M:%S")
    ws6["A3"] = "Core definitions"
    ws6["A4"] = "Total = all tracking numbers in the route"
    ws6["A5"] = "New = Status == NEW"
    ws6["A6"] = "OFD Total = Total - New"
    ws6["A7"] = "Success(Delivered) = Status == DELIVERED"
    ws6["A8"] = "Failed(*FAIL*) = Status contains FAIL"
    ws6["A9"] = "Remaining(OFD) = Status == OUT_FOR_DELIVERY"
    ws6["A10"] = "OFD(CompletionRate) = Success(Delivered) / OFD Total"
    ws6["A12"] = "Color rules"
    ws6["A13"] = "Purple: NO_DELIVERED"
    ws6["A14"] = "Yellow: MinutesSinceLast > 30 and <= 60"
    ws6["A15"] = "Red: MinutesSinceLast > 60"
    ws6["A17"] = "Exceptions criteria"
    ws6["A18"] = "1) NO_DELIVERED OR 2) MinutesSinceLast>120 & Remaining(OFD)>0 OR 3) DeliveriesPerHour<10 & Remaining(OFD)>0"
    ws6["A20"] = "3pm rule"
    ws6["A21"] = "At/after 3:00 PM ET: OFD(CompletionRate) < 50%"
    ws6["A23"] = "6pm rule"
    ws6["A24"] = "At/after 6:00 PM ET: OFD(CompletionRate) < 80%"
    autosize(ws6, cap=90)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), route_df, summary_df, exc_df


st.set_page_config(page_title="Daily DSP Operation Check", layout="wide")
st.title("Daily DSP Operation Check — 每天一键生成 Excel")

st.markdown(
    """
上传原始 Excel 后，系统会自动生成：
- **RouteMonitor**
- **Summary**
- **Exceptions**
- **3pm check**
- **6pm check**
- **Meta**
"""
)

with st.expander("输入要求", expanded=True):
    st.markdown(
        """
原始 Excel 必须满足：
- **B列** = Route
- **J列** = 包裹状态 Status
- **L列** = 状态改变时间

可选：
- 列名包含 **Driver** → 自动识别为 DriverName
- 列名包含 **Flee** → 自动识别为 FleeName
        """
    )

uploaded = st.file_uploader("📤 Upload your .xlsx file here", type=["xlsx"])

if uploaded:
    try:
        raw_df = pd.read_excel(uploaded, engine="openpyxl", dtype=str)
        output_bytes, route_df, summary_df, exc_df = build_excel_bytes(raw_df)

        total_routes = len(route_df)
        total_pkgs = int(route_df["Total"].sum()) if total_routes else 0
        ofd_total = int(route_df["OFD Total"].sum()) if total_routes else 0
        delivered = int(route_df["Success(Delivered)"].sum()) if total_routes else 0
        remaining_ofd = int(route_df["Remaining(OFD)"].sum()) if total_routes else 0
        new_pkgs = int(route_df["New"].sum()) if total_routes else 0
        ofd_rate = delivered / ofd_total if ofd_total > 0 else 0

        st.success("✅ 报表已生成")

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("Routes", total_routes)
        c2.metric("Total", total_pkgs)
        c3.metric("OFD Total", ofd_total)
        c4.metric("Delivered", delivered)
        c5.metric("Remaining(OFD)", remaining_ofd)
        c6.metric("OFD Rate", f"{ofd_rate:.2%}")

        st.caption(f"New packages: {new_pkgs} | Exceptions: {len(exc_df)}")

        ts = datetime.datetime.now(pytz.timezone(TZ)).strftime("%Y%m%d_%H%M%S")
        st.download_button(
            label="⬇️ Download Excel Report",
            data=output_bytes,
            file_name=f"route_monitor_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        with st.expander("Preview: RouteMonitor", expanded=False):
            st.dataframe(route_df, use_container_width=True)

        with st.expander("Preview: Summary", expanded=False):
            st.dataframe(summary_df, use_container_width=True)

        with st.expander("Preview: Exceptions", expanded=False):
            st.dataframe(exc_df, use_container_width=True)

    except Exception as e:
        st.error("生成失败，请检查文件格式是否正确。")
        st.exception(e)
else:
    st.info("请先上传 Excel 文件。")
