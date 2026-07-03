from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, size=12)
TITLE_FONT = Font(bold=True, size=14)
THIN_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

CATEGORY_ORDER = [
    "Early News",
    "Noon News",
    "Evening News",
    "Late News",
    "Prime News",
    "Liked Access",
    "Daytime",
]
CATEGORY_FILL = {
    "Early News": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Noon News": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Evening News": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Late News": PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid"),
    "Prime News": PatternFill(start_color="E6DCF1", end_color="E6DCF1", fill_type="solid"),
    "Liked Access": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "Daytime": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
}
DAY_MARK_FILL = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
DAY_COLS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
DAY_HEADER_LABELS = ["M", "T", "W", "Th", "F", "Sa", "Su"]


def _min_to_clock(m):
    h, mm = divmod(int(m) % (24 * 60), 60)
    suffix = "a" if h < 12 else "p"
    h12 = h % 12
    if h12 == 0:
        h12 = 12
    return f"{h12}:{mm:02d}{suffix}"


def _style_header(ws, row, ncols, start_col=1):
    for c in range(start_col, start_col + ncols):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _category_sort_key(cat):
    return CATEGORY_ORDER.index(cat) if cat in CATEGORY_ORDER else 99


def _group_spots_into_rows(spots):
    """Collapses individual spots into one flowchart row per unique
    (station, category, program, time), tracking how many spots land on
    each weekday -- long-format programs can carry more than one."""
    groups = defaultdict(lambda: {"day_counts": defaultdict(int), "rate": 0, "rating": 0})
    for s in spots:
        key = (s["category"], s["station"], s["program_name"], s["start_min"], s["end_min"])
        g = groups[key]
        g["day_counts"][s["day"]] += 1
        g["rate"] = s["rate"]
        g["rating"] = s["rating"]

    rows = []
    for (category, station, program, start_min, end_min), g in groups.items():
        spots_per_week = sum(g["day_counts"].values())
        rows.append(
            {
                "category": category,
                "station": station,
                "program": program,
                "start_min": start_min,
                "end_min": end_min,
                "day_counts": g["day_counts"],
                "rate": g["rate"],
                "rating": g["rating"],
                "spots_per_week": spots_per_week,
                "weekly_cost": g["rate"] * spots_per_week,
                "weekly_grps": g["rating"] * spots_per_week,
                "cpp": g["rate"] / g["rating"] if g["rating"] else 0,
            }
        )
    rows.sort(
        key=lambda r: (_category_sort_key(r["category"]), r["station"], r["start_min"])
    )
    return rows


def _write_station_summary_table(ws, result, start_row, ncols):
    """Compact 'totals by station, at a glance' block, sitting above the
    main flowchart grid so it doesn't disturb the timeslot-by-timeslot
    layout of the grid itself."""
    by_station = defaultdict(lambda: {"spots": 0, "grps": 0.0, "cost": 0.0})
    for s in result.spots:
        b = by_station[s["station"]]
        b["spots"] += 1
        b["grps"] += s["rating"]
        b["cost"] += s["rate"]

    ws.cell(row=start_row, column=1, value="Totals by Station").font = SECTION_FONT
    header_row = start_row + 1
    headers = ["Station", "Spots/Wk", "Wkly $", "Wkly GRPs", "CPP", "% Mkt GRPs"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    _style_header(ws, header_row, len(headers))

    r = header_row + 1
    for station in sorted(by_station):
        b = by_station[station]
        cpp = b["cost"] / b["grps"] if b["grps"] else 0
        pct = (b["grps"] / result.achieved_grps * 100) if result.achieved_grps else 0
        ws.cell(row=r, column=1, value=station)
        ws.cell(row=r, column=2, value=b["spots"])
        ws.cell(row=r, column=3, value=round(b["cost"], 0))
        ws.cell(row=r, column=4, value=round(b["grps"], 1))
        ws.cell(row=r, column=5, value=round(cpp, 2))
        ws.cell(row=r, column=6, value=round(pct, 1))
        r += 1

    ws.cell(row=r, column=1, value="MARKET").font = Font(bold=True)
    ws.cell(row=r, column=2, value=len(result.spots)).font = Font(bold=True)
    ws.cell(row=r, column=3, value=round(result.total_cost, 0)).font = Font(bold=True)
    ws.cell(row=r, column=4, value=round(result.achieved_grps, 1)).font = Font(bold=True)
    blended_cpp = result.total_cost / result.achieved_grps if result.achieved_grps else 0
    ws.cell(row=r, column=5, value=round(blended_cpp, 2)).font = Font(bold=True)
    ws.cell(row=r, column=6, value=100.0).font = Font(bold=True)

    return r + 2  # next free row, with one blank row of padding


def _write_flowchart_sheet(ws, result, target_demo_label):
    ws.sheet_view.showGridLines = False

    # Category, Station, Program, Time | 7 days | Spots,Rate,Rating,CPP,Wkly$,WklyGRPs,%MktGRPs
    ncols = 4 + len(DAY_COLS) + 7

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    ws.cell(row=1, column=1, value="Sample Weekly Buy Flowchart").font = TITLE_FONT

    blended_cpp = result.total_cost / result.achieved_grps if result.achieved_grps else 0
    subtitle = (
        f"Target demo: {target_demo_label}   |   "
        f"Weekly GRPs: {result.achieved_grps:.1f} of {result.target_grps:.0f} goal   |   "
        f"Weekly cost: ${result.total_cost:,.0f}   |   Blended CPP: ${blended_cpp:,.2f}"
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    ws.cell(row=2, column=1, value=subtitle).font = Font(italic=True, size=10)

    next_row = _write_station_summary_table(ws, result, start_row=4, ncols=ncols)

    header_row = next_row
    headers = (
        ["Category", "Station", "Program", "Time"]
        + DAY_HEADER_LABELS
        + ["Spots/Wk", "Rate", "Rating", "CPP", "Wkly $", "Wkly GRPs", "% Mkt GRPs"]
    )
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    _style_header(ws, header_row, len(headers))

    rows = _group_spots_into_rows(result.spots)
    base = 4 + len(DAY_COLS)  # last day column (Su)

    r = header_row + 1
    for row_data in rows:
        fill = CATEGORY_FILL.get(row_data["category"])
        c = 1
        for value in (row_data["category"], row_data["station"], row_data["program"]):
            cell = ws.cell(row=r, column=c, value=value)
            if fill:
                cell.fill = fill
            cell.border = THIN_BORDER
            c += 1
        time_label = f"{_min_to_clock(row_data['start_min'])}-{_min_to_clock(row_data['end_min'])}"
        cell = ws.cell(row=r, column=c, value=time_label)
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER
        c += 1

        for day in DAY_COLS:
            cell = ws.cell(row=r, column=c)
            cell.border = THIN_BORDER
            count = row_data["day_counts"].get(day, 0)
            if count == 1:
                cell.value = "●"
                cell.fill = DAY_MARK_FILL
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            elif count > 1:
                cell.value = count
                cell.fill = DAY_MARK_FILL
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            elif fill:
                cell.fill = fill
            c += 1

        tail_values = [
            row_data["spots_per_week"],
            row_data["rate"],
            row_data["rating"],
            round(row_data["cpp"], 2),
            round(row_data["weekly_cost"], 0),
            round(row_data["weekly_grps"], 1),
            None,
        ]
        for value in tail_values:
            cell = ws.cell(row=r, column=c, value=value)
            if fill:
                cell.fill = fill
            cell.border = THIN_BORDER
            c += 1
        r += 1

    total_row = r
    ws.cell(row=total_row, column=1, value="MARKET TOTAL").font = Font(bold=True)
    ws.cell(row=total_row, column=base + 1, value=sum(rd["spots_per_week"] for rd in rows)).font = Font(bold=True)
    ws.cell(row=total_row, column=base + 5, value=round(result.total_cost, 0)).font = Font(bold=True)
    ws.cell(row=total_row, column=base + 6, value=round(result.achieved_grps, 1)).font = Font(bold=True)
    ws.cell(row=total_row, column=base + 7, value=100.0).font = Font(bold=True)

    ws.freeze_panes = ws.cell(row=header_row + 1, column=4)

    widths = [13, 8, 36, 15] + [4] * len(DAY_COLS) + [9, 8, 8, 8, 11, 11, 11]
    _autofit(ws, widths)
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 14

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)
    ws.print_title_rows = f"{header_row}:{header_row}"


def write_workbook(result, path, target_demo_label="Adults 35+"):
    wb = Workbook()

    # ---- Buy Flowchart: the single-page, at-a-glance view ----
    ws = wb.active
    ws.title = "Buy Flowchart"
    _write_flowchart_sheet(ws, result, target_demo_label)

    # ---- Market Summary ----
    ws = wb.create_sheet("Market Summary")
    ws["A1"] = "Sample Weekly Buy -- Market Summary"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Target demo: {target_demo_label}"
    ws["A3"] = f"Target weekly GRPs: {result.target_grps:.0f}"
    ws["A4"] = f"Achieved weekly GRPs: {result.achieved_grps:.1f}"
    ws["A5"] = f"Total weekly cost: ${result.total_cost:,.0f}"
    blended_cpp = result.total_cost / result.achieved_grps if result.achieved_grps else 0
    ws["A6"] = f"Blended market CPP: ${blended_cpp:,.2f}"
    ws["A7"] = f"Total spots/week: {len(result.spots)}"

    row = 9
    if result.warnings:
        ws.cell(row=row, column=1, value="Notes").font = SECTION_FONT
        row += 1
        for w in result.warnings:
            ws.cell(row=row, column=1, value=f"- {w}")
            row += 1
        row += 1

    if result.outlier_avails:
        ws.cell(row=row, column=1, value="News avails deprioritized as CPP outliers").font = SECTION_FONT
        row += 1
        headers = ["Station", "Program", "CPP"]
        for c, h in enumerate(headers, start=1):
            ws.cell(row=row, column=c, value=h)
        _style_header(ws, row, len(headers))
        row += 1
        for station, program, cpp in result.outlier_avails:
            ws.cell(row=row, column=1, value=station)
            ws.cell(row=row, column=2, value=program)
            ws.cell(row=row, column=3, value=cpp)
            row += 1
        row += 1

    # breakdown by category
    ws.cell(row=row, column=1, value="Breakdown by Daypart Category").font = SECTION_FONT
    row += 1
    headers = ["Category", "Spots", "Weekly GRPs", "Weekly Cost", "CPP"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=row, column=c, value=h)
    _style_header(ws, row, len(headers))
    row += 1
    by_cat = defaultdict(lambda: {"spots": 0, "grps": 0.0, "cost": 0.0})
    for s in result.spots:
        b = by_cat[s["category"]]
        b["spots"] += 1
        b["grps"] += s["rating"]
        b["cost"] += s["rate"]
    for cat in CATEGORY_ORDER:
        if cat not in by_cat:
            continue
        b = by_cat[cat]
        cpp = b["cost"] / b["grps"] if b["grps"] else 0
        ws.cell(row=row, column=1, value=cat)
        ws.cell(row=row, column=2, value=b["spots"])
        ws.cell(row=row, column=3, value=round(b["grps"], 1))
        ws.cell(row=row, column=4, value=round(b["cost"], 0))
        ws.cell(row=row, column=5, value=round(cpp, 2))
        row += 1
    row += 1

    # breakdown by station
    ws.cell(row=row, column=1, value="Breakdown by Station").font = SECTION_FONT
    row += 1
    headers = ["Station", "Spots", "Weekly GRPs", "Weekly Cost", "CPP"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=row, column=c, value=h)
    _style_header(ws, row, len(headers))
    row += 1
    by_station = defaultdict(lambda: {"spots": 0, "grps": 0.0, "cost": 0.0})
    for s in result.spots:
        b = by_station[s["station"]]
        b["spots"] += 1
        b["grps"] += s["rating"]
        b["cost"] += s["rate"]
    for station in sorted(by_station):
        b = by_station[station]
        cpp = b["cost"] / b["grps"] if b["grps"] else 0
        ws.cell(row=row, column=1, value=station)
        ws.cell(row=row, column=2, value=b["spots"])
        ws.cell(row=row, column=3, value=round(b["grps"], 1))
        ws.cell(row=row, column=4, value=round(b["cost"], 0))
        ws.cell(row=row, column=5, value=round(cpp, 2))
        row += 1

    _autofit(ws, [42, 12, 14, 14, 10])

    # ---- Per-station tabs ----
    stations = sorted({s["station"] for s in result.spots})
    day_order = {d: i for i, d in enumerate(DAY_COLS)}

    for station in stations:
        ws = wb.create_sheet(station)
        headers = ["Category", "Program", "Day", "Time", "Rate", "Rating", "CPP"]
        for c, h in enumerate(headers, start=1):
            ws.cell(row=1, column=c, value=h)
        _style_header(ws, 1, len(headers))

        station_spots = [s for s in result.spots if s["station"] == station]
        station_spots.sort(
            key=lambda s: (
                _category_sort_key(s["category"]),
                s["program_name"],
                day_order.get(s["day"], 9),
            )
        )
        r = 2
        for s in station_spots:
            ws.cell(row=r, column=1, value=s["category"])
            ws.cell(row=r, column=2, value=s["program_name"])
            ws.cell(row=r, column=3, value=s["day"])
            ws.cell(row=r, column=4, value=_min_to_clock(s["start_min"]))
            ws.cell(row=r, column=5, value=s["rate"])
            ws.cell(row=r, column=6, value=s["rating"])
            ws.cell(row=r, column=7, value=s["cpp"])
            r += 1

        total_grps = sum(s["rating"] for s in station_spots)
        total_cost = sum(s["rate"] for s in station_spots)
        ws.cell(row=r + 1, column=1, value="TOTAL").font = Font(bold=True)
        ws.cell(row=r + 1, column=3, value=f"{len(station_spots)} spots")
        ws.cell(row=r + 1, column=5, value=round(total_cost, 0)).font = Font(bold=True)
        ws.cell(row=r + 1, column=6, value=round(total_grps, 1)).font = Font(bold=True)
        ws.cell(row=r + 1, column=7, value=round(total_cost / total_grps, 2) if total_grps else 0).font = Font(bold=True)

        _autofit(ws, [14, 38, 8, 10, 10, 10, 10])

    # ---- All Spots flat tab ----
    ws = wb.create_sheet("All Spots")
    headers = ["Station", "Category", "Program", "Day", "Time", "Rate", "Rating", "CPP"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)
    _style_header(ws, 1, len(headers))
    all_sorted = sorted(
        result.spots,
        key=lambda s: (s["station"], day_order.get(s["day"], 9), s["start_min"]),
    )
    r = 2
    for s in all_sorted:
        ws.cell(row=r, column=1, value=s["station"])
        ws.cell(row=r, column=2, value=s["category"])
        ws.cell(row=r, column=3, value=s["program_name"])
        ws.cell(row=r, column=4, value=s["day"])
        ws.cell(row=r, column=5, value=_min_to_clock(s["start_min"]))
        ws.cell(row=r, column=6, value=s["rate"])
        ws.cell(row=r, column=7, value=s["rating"])
        ws.cell(row=r, column=8, value=s["cpp"])
        r += 1
    _autofit(ws, [10, 14, 38, 8, 10, 10, 10, 10])

    wb.save(path)
