from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SECTION_FONT = Font(bold=True, size=12)
CATEGORY_ORDER = ["Early News", "Noon News", "Evening News", "Late News", "Liked Access", "Daytime"]


def _min_to_clock(m):
    h, mm = divmod(int(m) % (24 * 60), 60)
    suffix = "a" if h < 12 else "p"
    h12 = h % 12
    if h12 == 0:
        h12 = 12
    return f"{h12}:{mm:02d}{suffix}"


def _style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")


def _autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_workbook(result, path, target_demo_label="Adults 35+"):
    wb = Workbook()

    # ---- Market Summary ----
    ws = wb.active
    ws.title = "Market Summary"
    ws["A1"] = "Sample Weekly Buy -- Market Summary"
    ws["A1"].font = Font(bold=True, size=14)
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
    day_order = {d: i for i, d in enumerate(["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])}

    for station in stations:
        ws = wb.create_sheet(station)
        headers = ["Category", "Program", "Day", "Time", "Rate", f"Rating", "CPP"]
        for c, h in enumerate(headers, start=1):
            ws.cell(row=1, column=c, value=h)
        _style_header(ws, 1, len(headers))

        station_spots = [s for s in result.spots if s["station"] == station]
        station_spots.sort(
            key=lambda s: (
                CATEGORY_ORDER.index(s["category"]) if s["category"] in CATEGORY_ORDER else 99,
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
