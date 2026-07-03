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

# Fixed column layout for the flowchart grid.
COL_CATEGORY = 1
COL_STATION = 2
COL_PROGRAM = 3
COL_TIME = 4
COL_DAY_FIRST = 5
COL_DAY_LAST = COL_DAY_FIRST + len(DAY_COLS) - 1  # 11 (Sun)
COL_SPOTS = COL_DAY_LAST + 1  # 12
COL_RATE = COL_SPOTS + 1  # 13
COL_RATING = COL_RATE + 1  # 14
COL_CPP = COL_RATING + 1  # 15
COL_WKLY_COST = COL_CPP + 1  # 16
COL_WKLY_GRPS = COL_WKLY_COST + 1  # 17
COL_PCT_MKT = COL_WKLY_GRPS + 1  # 18
NCOLS = COL_PCT_MKT


def _cl(col):
    return get_column_letter(col)


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


def _group_avails_into_rows(eligible_avails, spots):
    """One flowchart row per unique (category, station, program, time),
    covering every eligible avail -- not just the ones actually bought.
    Days the program doesn't air are omitted; days it airs but wasn't
    bought get an explicit 0 so the buy can be edited in place."""
    identities = {}
    for a in eligible_avails:
        key = (a._category, a.station, a.program_name, a.start_min, a.end_min)
        info = identities.setdefault(key, {"days": set(), "rate": a.rate, "rating": a.rating})
        info["days"].update(a.days)

    bought_counts = defaultdict(lambda: defaultdict(int))
    for s in spots:
        key = (s["category"], s["station"], s["program_name"], s["start_min"], s["end_min"])
        bought_counts[key][s["day"]] += 1

    rows = []
    for (category, station, program, start_min, end_min), info in identities.items():
        counts = bought_counts.get((category, station, program, start_min, end_min), {})
        day_counts = {d: counts.get(d, 0) for d in info["days"]}
        rows.append(
            {
                "category": category,
                "station": station,
                "program": program,
                "start_min": start_min,
                "end_min": end_min,
                "day_counts": day_counts,
                "rate": info["rate"],
                "rating": info["rating"],
            }
        )
    rows.sort(key=lambda r: (r["station"], r["program"], r["start_min"]))
    return rows


def _write_station_summary_table(ws, result, start_row, stations, grid_first_row, grid_last_row, grid_total_row):
    """Compact 'totals by station, at a glance' block above the main grid.
    Every figure is a live formula pulling from the grid below, so it
    stays correct if the grid's spot counts are edited."""
    station_col = _cl(COL_STATION)
    spots_col = _cl(COL_SPOTS)
    cost_col = _cl(COL_WKLY_COST)
    grps_col = _cl(COL_WKLY_GRPS)

    ws.cell(row=start_row, column=1, value="Totals by Station (live -- recalculates if you edit the grid below)").font = SECTION_FONT
    header_row = start_row + 1
    headers = ["Station", "Spots/Wk", "Wkly $", "Wkly GRPs", "CPP", "% Mkt GRPs"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    _style_header(ws, header_row, len(headers))

    r = header_row + 1
    for station in stations:
        rng = lambda col: f"{col}{grid_first_row}:{col}{grid_last_row}"
        ws.cell(row=r, column=1, value=station)
        ws.cell(
            row=r, column=2,
            value=f'=SUMIF({rng(station_col)},"{station}",{rng(spots_col)})',
        )
        ws.cell(
            row=r, column=3,
            value=f'=SUMIF({rng(station_col)},"{station}",{rng(cost_col)})',
        )
        ws.cell(
            row=r, column=4,
            value=f'=SUMIF({rng(station_col)},"{station}",{rng(grps_col)})',
        )
        ws.cell(row=r, column=5, value=f"=IFERROR(C{r}/D{r},0)")
        ws.cell(
            row=r, column=6,
            value=f"=IFERROR(D{r}/{grps_col}${grid_total_row}*100,0)",
        )
        r += 1

    market_row = r
    ws.cell(row=market_row, column=1, value="MARKET").font = Font(bold=True)
    for col, src_col in ((2, _cl(COL_SPOTS)), (3, _cl(COL_WKLY_COST)), (4, _cl(COL_WKLY_GRPS))):
        cell = ws.cell(row=market_row, column=col, value=f"={src_col}{grid_total_row}")
        cell.font = Font(bold=True)
    ws.cell(row=market_row, column=5, value=f"=IFERROR(C{market_row}/D{market_row},0)").font = Font(bold=True)
    ws.cell(row=market_row, column=6, value=f"=IFERROR(D{market_row}/{grps_col}${grid_total_row}*100,0)").font = Font(bold=True)

    return market_row + 2  # next free row, with a blank row of padding


def _write_flowchart_sheet(ws, result, target_demo_label):
    ws.sheet_view.showGridLines = False

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=NCOLS)
    ws.cell(row=1, column=1, value="Sample Weekly Buy Flowchart").font = TITLE_FONT

    blended_cpp = result.total_cost / result.achieved_grps if result.achieved_grps else 0
    subtitle = (
        f"Target demo: {target_demo_label}   |   "
        f"Weekly GRPs: {result.achieved_grps:.1f} of {result.target_grps:.0f} goal   |   "
        f"Weekly cost: ${result.total_cost:,.0f}   |   Blended CPP: ${blended_cpp:,.2f}   |   "
        f"Edit the day columns below to change the buy -- totals recalculate automatically."
    )
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=NCOLS)
    ws.cell(row=2, column=1, value=subtitle).font = Font(italic=True, size=10)

    rows = _group_avails_into_rows(result.eligible_avails, result.spots)
    stations = sorted({rd["station"] for rd in rows})

    # Layout is fully determined up front so the station-summary formulas
    # (written first, but referencing the grid below) point at the right rows.
    summary_start = 4
    summary_header_row = summary_start + 1
    summary_market_row = summary_header_row + 1 + len(stations)
    header_row = summary_market_row + 2
    grid_first_row = header_row + 1
    grid_last_row = grid_first_row + len(rows) - 1
    total_row = grid_last_row + 1

    _write_station_summary_table(
        ws, result, summary_start, stations, grid_first_row, grid_last_row, total_row
    )

    headers = (
        ["Category", "Station", "Program", "Time"]
        + DAY_HEADER_LABELS
        + ["Spots/Wk", "Rate", "Rating", "CPP", "Wkly $", "Wkly GRPs", "% Mkt GRPs"]
    )
    for c, h in enumerate(headers, start=1):
        ws.cell(row=header_row, column=c, value=h)
    _style_header(ws, header_row, len(headers))

    day_first_letter = _cl(COL_DAY_FIRST)
    day_last_letter = _cl(COL_DAY_LAST)
    rate_letter = _cl(COL_RATE)
    rating_letter = _cl(COL_RATING)
    spots_letter = _cl(COL_SPOTS)
    grps_letter = _cl(COL_WKLY_GRPS)

    r = grid_first_row
    for row_data in rows:
        fill = CATEGORY_FILL.get(row_data["category"])
        for col, value in (
            (COL_CATEGORY, row_data["category"]),
            (COL_STATION, row_data["station"]),
            (COL_PROGRAM, row_data["program"]),
        ):
            cell = ws.cell(row=r, column=col, value=value)
            if fill:
                cell.fill = fill
            cell.border = THIN_BORDER

        time_label = f"{_min_to_clock(row_data['start_min'])}-{_min_to_clock(row_data['end_min'])}"
        cell = ws.cell(row=r, column=COL_TIME, value=time_label)
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        for i, day in enumerate(DAY_COLS):
            col = COL_DAY_FIRST + i
            cell = ws.cell(row=r, column=col)
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center")
            if day in row_data["day_counts"]:
                count = row_data["day_counts"][day]
                cell.value = count
                if count > 0:
                    cell.fill = DAY_MARK_FILL
                    cell.font = Font(color="FFFFFF", bold=True)
                else:
                    if fill:
                        cell.fill = fill
                    cell.font = Font(color="595959")
            elif fill:
                cell.fill = fill

        cell = ws.cell(row=r, column=COL_SPOTS, value=f"=SUM({day_first_letter}{r}:{day_last_letter}{r})")
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=r, column=COL_RATE, value=row_data["rate"])
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=r, column=COL_RATING, value=row_data["rating"])
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=r, column=COL_CPP, value=f"=IFERROR({rate_letter}{r}/{rating_letter}{r},0)")
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=r, column=COL_WKLY_COST, value=f"={spots_letter}{r}*{rate_letter}{r}")
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(row=r, column=COL_WKLY_GRPS, value=f"={spots_letter}{r}*{rating_letter}{r}")
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        cell = ws.cell(
            row=r, column=COL_PCT_MKT,
            value=f"=IFERROR({grps_letter}{r}/{grps_letter}${total_row}*100,0)",
        )
        if fill:
            cell.fill = fill
        cell.border = THIN_BORDER

        r += 1

    ws.cell(row=total_row, column=COL_CATEGORY, value="MARKET TOTAL").font = Font(bold=True)
    ws.cell(
        row=total_row, column=COL_SPOTS,
        value=f"=SUM({spots_letter}{grid_first_row}:{spots_letter}{grid_last_row})",
    ).font = Font(bold=True)
    ws.cell(
        row=total_row, column=COL_WKLY_COST,
        value=f"=SUM({_cl(COL_WKLY_COST)}{grid_first_row}:{_cl(COL_WKLY_COST)}{grid_last_row})",
    ).font = Font(bold=True)
    ws.cell(
        row=total_row, column=COL_WKLY_GRPS,
        value=f"=SUM({grps_letter}{grid_first_row}:{grps_letter}{grid_last_row})",
    ).font = Font(bold=True)
    ws.cell(
        row=total_row, column=COL_CPP,
        value=f"=IFERROR({_cl(COL_WKLY_COST)}{total_row}/{grps_letter}{total_row},0)",
    ).font = Font(bold=True)
    ws.cell(
        row=total_row, column=COL_PCT_MKT,
        value=f"=IFERROR({grps_letter}{total_row}/{grps_letter}{total_row}*100,0)",
    ).font = Font(bold=True)

    ws.freeze_panes = ws.cell(row=header_row + 1, column=COL_DAY_FIRST)

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

    # ---- Buy Flowchart: the single-page, at-a-glance, editable view ----
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
