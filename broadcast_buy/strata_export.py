"""Exports a sample buy as a Strata-importable order (the "ADX" XML format
Strata itself exports as .sbx). Reverse-engineered from a real buy export
rather than public documentation, so some fields Strata didn't need for
that import (market/station numeric codes, agency contact details) are
left blank for the buyer to fill in after import.

Key discovery: setting <buyType>daily</buyType> makes each <week> entry
represent one calendar day (not a calendar week) and each line's
<spot><quantity> the exact spot count for that date -- which maps directly
onto our own day-by-day model with no ambiguity about which specific days
get the spots.
"""

import datetime
import xml.etree.ElementTree as ET

from .grouping import group_avails_into_rows

DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
STRATA_DAY_TAGS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def _sub(parent, tag, text=None, **attrib):
    el = ET.SubElement(parent, tag, {k: str(v) for k, v in attrib.items() if v is not None})
    if text is not None:
        el.text = str(text)
    return el


def _spot_length_to_duration(spot_length):
    """'00:00:30' -> 'PT30S' (ISO 8601 duration, seconds only -- matches
    what real Strata exports use for standard :15/:30/:60 spot lengths).
    Also handles the parser's bare ':30' fallback format."""
    parts = [int(p) for p in spot_length.split(":") if p]
    while len(parts) < 3:
        parts.insert(0, 0)
    h, m, s = parts
    return f"PT{h * 3600 + m * 60 + s}S"


def _min_to_hms(minutes):
    minutes = int(minutes) % 1440
    h, m = divmod(minutes, 60)
    return f"{h:02d}:{m:02d}:00"


def _parse_date(value):
    if isinstance(value, datetime.date):
        return value
    return datetime.date.fromisoformat(value)


def _daterange(start, end):
    dates = []
    d = start
    while d <= end:
        dates.append(d)
        d += datetime.timedelta(days=1)
    return dates


def _row_total_spots(row, dates):
    return sum(row["day_counts"].get(DAY_NAMES[dt.weekday()], 0) for dt in dates)


def write_strata_order(
    result,
    output_path,
    *,
    market_name,
    flight_start,
    flight_end,
    target_demo_group="Adults",
    target_demo_age=35,
    campaign_name=None,
    agency_name="GPS Impact",
):
    """Writes a Strata ADX order that repeats the sample buy's weekly
    pattern across every day of [flight_start, flight_end] (inclusive).
    Only rows with at least one bought spot are included -- the sample
    buy's "shown but never bought" rows (Prime, unpicked Daytime, etc.)
    have no place in an actual station order."""
    start = _parse_date(flight_start)
    end = _parse_date(flight_end)
    if end < start:
        raise ValueError("flight_end must be on or after flight_start")
    dates = _daterange(start, end)
    campaign_name = campaign_name or "Sample Buy"

    rows = group_avails_into_rows(result.eligible_avails, result.spots)
    bought_rows = [r for r in rows if sum(r["day_counts"].values()) > 0]

    root = ET.Element(
        "adx",
        {"xsi:noNamespaceSchemaLocation": "", "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"},
    )

    document = _sub(root, "document")
    _sub(document, "schemaVersion", "1.0")
    _sub(document, "name", campaign_name)
    _sub(document, "date", datetime.date.today().isoformat())
    _sub(document, "mediaType", "Spotcable")
    _sub(document, "documentType")
    _sub(document, "documentCode")

    campaign = _sub(root, "campaign")
    key = _sub(campaign, "key", codeOwner="NCC", codeDescription="CampaignID")
    _sub(key, "id")
    key = _sub(campaign, "key", codeOwner="Strata", codeDescription="DMA Override")
    _sub(key, "id", "0")
    key = _sub(campaign, "key", codeOwner="Strata", codeDescription="Zone Pops")
    _sub(key, "id", "max")
    key = _sub(campaign, "key", codeOwner="VIEW32", codeDescription="CampaignName")
    _sub(key, "id", campaign_name)

    date_range = _sub(campaign, "dateRange")
    _sub(date_range, "startDate", start.isoformat())
    _sub(date_range, "endDate", end.isoformat())

    rep = _sub(campaign, "company", type="Rep")
    _sub(rep, "name")
    office = _sub(rep, "office")
    _sub(office, "name")
    address = _sub(office, "address")
    _sub(address, "street")
    _sub(address, "city")
    _sub(address, "state", code="__")
    _sub(address, "postalCode", "00000")
    _sub(office, "phone", type="voice")
    rep_contact = _sub(rep, "contact", role="AE")
    _sub(rep_contact, "firstName")
    _sub(rep_contact, "lastName")
    _sub(rep_contact, "email")
    _sub(rep_contact, "phone", type="voice")

    agency = _sub(campaign, "company", type="Agency")
    _sub(agency, "name", agency_name)
    agency_contact = _sub(agency, "contact", role="Buyer")
    _sub(agency_contact, "firstName")
    _sub(agency_contact, "lastName")
    _sub(agency_contact, "email")
    _sub(agency_contact, "phone", type="voice")
    agency_id = _sub(agency, "ID")
    _sub(agency_id, "code", codeOwner="Agency")

    advertiser = _sub(campaign, "advertiser")
    _sub(advertiser, "name")
    adv_id = _sub(advertiser, "ID")
    _sub(adv_id, "code", codeOwner="Agency")

    product = _sub(campaign, "product")
    _sub(product, "name")
    prod_id = _sub(product, "ID")
    _sub(prod_id, "code", codeOwner="Agency")

    estimate = _sub(campaign, "estimate")
    _sub(estimate, "desc")
    est_id = _sub(estimate, "ID")
    _sub(est_id, "code", codeOwner="Agency")

    makegood = _sub(campaign, "makeGoodPolicy")
    _sub(makegood, "code")

    demo = _sub(campaign, "demo", demoRank="1")
    _sub(demo, "group", target_demo_group)
    _sub(demo, "ageFrom", target_demo_age)
    _sub(demo, "ageTo", 99)

    _sub(campaign, "buyType", "daily")
    _sub(campaign, "populations", "0", demoRank="1")

    row_spots = {id(r): _row_total_spots(r, dates) for r in bought_rows}
    total_spots = sum(row_spots.values())
    total_cost = sum(row_spots[id(r)] * r["rate"] for r in bought_rows)

    order = _sub(campaign, "order")
    key = _sub(order, "key", codeOwner="NCC", codeDescription="Market")
    _sub(key, "id")  # left blank -- no Strata numeric market code available from rate cards
    key = _sub(order, "key", codeOwner="")
    _sub(key, "id")
    key = _sub(order, "key", codeOwner="Strata", codeDescription="Pops")
    _sub(key, "id", "dma;book")
    key = _sub(order, "key", codeOwner="Strata", codeDescription="UseBroadcastWeeks")
    _sub(key, "id", "0")

    order_totals = _sub(order, "totals")
    _sub(order_totals, "cost", f"{total_cost:.2f}")
    _sub(order_totals, "spots", total_spots)

    market_el = _sub(order, "market")
    _sub(market_el, "name", market_name)

    survey = _sub(order, "survey")
    _sub(survey, "ratingService")
    _sub(survey, "geography")
    _sub(survey, "shareBook")
    _sub(survey, "PUTBook")
    _sub(survey, "profile")
    _sub(survey, "comment", codeOwner="Spotcable")

    _sub(order, "populations", "0", demoRank="1")
    _sub(order, "comment")

    system_order = _sub(order, "systemOrder")
    key = _sub(system_order, "key", codeOwner="")
    _sub(key, "id")
    key = _sub(system_order, "key", codeOwner="Strata", codeDescription="UseZonePop")
    _sub(key, "id", "false")
    _sub(system_order, "comment")
    system = _sub(system_order, "system")
    _sub(system, "name")
    _sub(system, "syscode")
    _sub(system_order, "affiliateSplit")
    _sub(system_order, "populations", "0", demoRank="1")

    so_totals = _sub(system_order, "totals")
    _sub(so_totals, "cost", f"{total_cost:.2f}")
    _sub(so_totals, "spots", total_spots)

    weeks = _sub(system_order, "weeks", count=len(dates))
    for i, dt in enumerate(dates, start=1):
        _sub(weeks, "week", number=i, startDate=dt.isoformat())

    for row in bought_rows:
        detail_line = _sub(system_order, "detailLine", detailLineID="0")
        _sub(detail_line, "startTime", _min_to_hms(row["start_min"]))
        _sub(detail_line, "endTime", _min_to_hms(row["end_min"]))

        bought_days = [d for d in DAY_NAMES if row["day_counts"].get(d, 0) > 0]
        _sub(detail_line, "startDay", bought_days[0][:2] if bought_days else DAY_NAMES[0][:2])

        day_of_week = _sub(detail_line, "dayOfWeek")
        for day_name, tag in zip(DAY_NAMES, STRATA_DAY_TAGS):
            _sub(day_of_week, tag, "Y" if day_name in bought_days else "N")

        _sub(detail_line, "length", _spot_length_to_duration(row["spot_length"]))
        _sub(detail_line, "daypartCode", row["daypart"])
        _sub(detail_line, "program", row["program"])

        network = _sub(detail_line, "network")
        _sub(network, "name", row["station"])
        net_id = _sub(network, "ID")
        _sub(net_id, "code", row["station"], codeOwner="Spotcable")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", row["station"], codeOwner="Strata", codeDescription="Station")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", "TV", codeOwner="Strata", codeDescription="Band")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", start.isoformat(), codeOwner="Strata", codeDescription="Start Date")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", end.isoformat(), codeOwner="Strata", codeDescription="End Date")

        _sub(detail_line, "spotCost", f"{row['rate']:.2f}", currency="USD")

        demo_value = _sub(detail_line, "demoValue", demoRank="1")
        _sub(demo_value, "value", row["rating"], type="Ratings")

        line_spots = row_spots[id(row)]
        line_totals = _sub(detail_line, "totals")
        _sub(line_totals, "cost", f"{line_spots * row['rate']:.2f}", currency="USD")
        _sub(line_totals, "spots", line_spots)

        for i, dt in enumerate(dates, start=1):
            spot = _sub(detail_line, "spot")
            _sub(spot, "weekNumber", i)
            _sub(spot, "quantity", row["day_counts"].get(DAY_NAMES[dt.weekday()], 0))

    tree = ET.ElementTree(root)
    ET.indent(tree, space="\t")
    tree.write(output_path, encoding="UTF-8", xml_declaration=True)
