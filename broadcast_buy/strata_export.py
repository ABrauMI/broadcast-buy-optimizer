"""Exports a sample buy as a Strata-importable order (the "ADX" XML format
Strata itself exports as .sbx). Reverse-engineered from real buy exports
rather than public documentation. Market and station numeric codes (Nielsen
market IDs, Spotcable station codes) come from a small lookup table below,
since Strata assigns those centrally and a rate card has no way to carry
them -- anything not in that table is left blank for the buyer to fill in
after import, along with agency contact details we simply don't track.

Key discovery: setting <buyType>daily</buyType> makes each <week> entry
represent one calendar day (not a calendar week) and each line's
<spot><quantity> the exact spot count for that date -- which maps directly
onto our own day-by-day model with no ambiguity about which specific days
get the spots.

The other hard-won discovery: <daypartCode> must be one of Strata's own
enumerated daypart labels ("Early Morning", "Early News", ...), not our
internal 2-letter Excel-sheet shorthand -- an unrecognized value is what
produced a JScript "Object required" error on import.
"""

import datetime
import xml.etree.ElementTree as ET

from .classify import daypart_code

DAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
STRATA_DAY_TAGS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

# Strata's daypartCode expects one of its own enumerated daypart labels, not
# our internal 2-letter Excel-sheet code (EM/EN/LN/...) -- feeding it the
# short code is what produced the "Object required" JScript error on import,
# since Strata's transform looks the value up in its own daypart table and
# gets nothing back for a code it's never heard of. Anything we can't map
# confidently falls back to "Unassigned", which is itself a value real
# Strata orders use.
STRATA_DAYPART_NAME = {
    "EM": "Early Morning",
    "EN": "Early News",
    "LN": "Late News",
    "DY": "Daytime",
    "PA": "Access",
    "PR": "Prime",
    "EF": "Early Fringe",
    "LF": "Late Fringe",
    "SP": "Sports",
}


def _strata_daypart(daypart_name):
    return STRATA_DAYPART_NAME.get(daypart_code(daypart_name), "Unassigned")


# Nielsen "Spotcable" station codes and Strata market codes/DMA populations
# for the stations and markets GPS Impact has actually bought through
# Strata, pulled from real Strata exports for those buys rather than
# anything derivable from a rate card -- Strata assigns these centrally, so
# a station's own avail file has no way to carry them. Extend both tables
# as new stations or markets come up; anything missing is left blank
# rather than guessed, since a wrong numeric code would point Strata at
# the wrong station/market outright.
STATION_INFO = {
    "WISN": {"spotcable_code": "5295", "network_name": "WISN-TV"},
    "WTMJ": {"spotcable_code": "5108", "network_name": "WTMJ-TV"},
    "WITI": {"spotcable_code": "5369", "network_name": "WITI-TV"},
    "WDJT": {"spotcable_code": "6378", "network_name": "WDJT-TV"},
    "WISC": {"spotcable_code": "5378", "network_name": "WISC-TV"},
    "WMTV": {"spotcable_code": "5741", "network_name": "WMTV-TV"},
    "WMSN": {"spotcable_code": "6294", "network_name": "WMSN-TV"},
    "WKOW": {"spotcable_code": "5736", "network_name": "WKOW-TV"},
    # Montana's 1st congressional district spans the Butte and Missoula
    # DMAs. Several of these are duopoly subchannels rather than
    # standalone stations, which is why Strata's own network name carries
    # a "+S2" (second subchannel) suffix instead of the usual "-TV".
    "KECI": {"spotcable_code": "312", "network_name": "KECI+S2"},
    "KPAX": {"spotcable_code": "532", "network_name": "KPAX+S2"},
    "KTMF": {"spotcable_code": "2022", "network_name": "KTMF+S2"},
    "KTVM": {"spotcable_code": "5530", "network_name": "KTVM-TV"},
    "KWYB": {"spotcable_code": "2402", "network_name": "KWYB+S2"},
    "KXLF": {"spotcable_code": "186", "network_name": "KXLF+S2"},
    "NTMF": {"spotcable_code": "1703", "network_name": "NTMF-TV"},
    "NWYB": {"spotcable_code": "1704", "network_name": "NWYB-TV"},
}

MARKET_INFO = {
    "milwaukee": {
        "strata_name": "Milwaukee-Racine",
        "nsi_id": "217",
        "ncc_market_id": "193",
        # DMA population by (demo group, age-from), as Strata itself reports it.
        "population": {
            ("ADULTS", 18): 1795080,
            ("ADULTS", 25): 1583822,
            ("ADULTS", 35): 1298545,
            ("ADULTS", 50): 863935,
            ("ADULTS", 55): 729670,
            ("ADULTS", 65): 439028,
            ("WOMEN", 35): 672671,
            ("HOUSEHOLDS", 0): 957240,
        },
    },
    "madison": {
        "strata_name": "Madison",
        "nsi_id": "269",
        "ncc_market_id": "180",
        "population": {
            ("ADULTS", 18): 826698,
            ("ADULTS", 25): 715936,
            ("ADULTS", 35): 578483,
            ("ADULTS", 50): 377973,
            ("ADULTS", 55): 318051,
            ("ADULTS", 65): 193633,
            ("WOMEN", 35): 293969,
            ("HOUSEHOLDS", 0): 441730,
        },
    },
    "butte": {
        "strata_name": "Butte, MT",
        "nsi_id": "354",
        "ncc_market_id": "393",
        "population": {
            ("ADULTS", 35): 112283,
            ("ADULTS", 50): 73312,
        },
    },
    "missoula": {
        "strata_name": "Missoula",
        "nsi_id": "362",
        "ncc_market_id": "404",
        "population": {
            ("ADULTS", 35): 199170,
            ("ADULTS", 50): 136799,
        },
    },
}


def _match_market(market_name):
    """Case-insensitive substring match against the known market table --
    lets 'Milwaukee, WI' or 'Milwaukee' both resolve to the same entry."""
    lowered = market_name.lower()
    for key, info in MARKET_INFO.items():
        if key in lowered:
            return info
    return None


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
    rows,
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
    rows is the same shape group_avails_into_rows() and
    workbook_import.read_sample_buy_rows() both produce -- callers decide
    whether that's freshly computed from a BuyResult or read back from a
    (possibly hand-edited) Sample Buy workbook; this function doesn't care
    which. Only rows with at least one bought spot are included -- the
    sample buy's "shown but never bought" rows (Prime, unpicked Daytime,
    etc.) have no place in an actual station order. Returns a list of
    warnings for anything (market, station) not in the known Strata code
    tables, since those fields get left blank rather than guessed at."""
    start = _parse_date(flight_start)
    end = _parse_date(flight_end)
    if end < start:
        raise ValueError("flight_end must be on or after flight_start")
    dates = _daterange(start, end)
    campaign_name = campaign_name or "Sample Buy"
    warnings = []

    market_info = _match_market(market_name)
    if market_info is None:
        warnings.append(
            f"Market {market_name!r} isn't in the known Strata market table -- "
            "market/DMA numeric codes left blank in the .sbx; fill them in after import."
        )
    strata_market_name = market_info["strata_name"] if market_info else market_name
    market_nsi_id = market_info["nsi_id"] if market_info else ""
    market_ncc_id = market_info["ncc_market_id"] if market_info else None
    population = market_info["population"].get((target_demo_group.upper(), target_demo_age)) if market_info else None
    if market_info and population is None:
        warnings.append(
            f"No DMA population on file for {target_demo_group} {target_demo_age}+ in {strata_market_name} -- "
            "Impressions left as 0 in the .sbx."
        )

    bought_rows = [r for r in rows if sum(r["day_counts"].values()) > 0]

    unmapped_stations = sorted({r["station"] for r in bought_rows if r["station"].upper() not in STATION_INFO})
    for station in unmapped_stations:
        warnings.append(
            f"No known Spotcable station code for {station!r} -- codeDescription left blank in the .sbx."
        )

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
    _sub(campaign, "populations", population if population is not None else "0", demoRank="1")

    row_spots = {id(r): _row_total_spots(r, dates) for r in bought_rows}
    total_spots = sum(row_spots.values())
    total_cost = sum(row_spots[id(r)] * r["rate"] for r in bought_rows)

    order = _sub(campaign, "order")
    key = _sub(order, "key", codeOwner="NCC", codeDescription="Market")
    _sub(key, "id", market_ncc_id)
    key = _sub(order, "key", codeOwner="")
    _sub(key, "id")
    key = _sub(order, "key", codeOwner="Strata", codeDescription="Pops")
    _sub(key, "id", "dma;book")
    key = _sub(order, "key", codeOwner="Strata", codeDescription="UseBroadcastWeeks")
    _sub(key, "id", "0")

    order_totals = _sub(order, "totals")
    _sub(order_totals, "cost", f"{total_cost:.2f}")
    _sub(order_totals, "spots", total_spots)

    market_el = _sub(order, "market", nsi_id=market_nsi_id)
    _sub(market_el, "name", strata_market_name)

    survey = _sub(order, "survey")
    _sub(survey, "ratingService")
    _sub(survey, "geography")
    _sub(survey, "shareBook")
    _sub(survey, "PUTBook")
    _sub(survey, "profile")
    _sub(survey, "comment", codeOwner="Spotcable")

    _sub(order, "populations", population if population is not None else "0", demoRank="1")
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
    _sub(system_order, "populations", population if population is not None else "0", demoRank="1")

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
        _sub(detail_line, "daypartCode", _strata_daypart(row["daypart_name"]))
        _sub(detail_line, "program", row["program"])
        _sub(detail_line, "comment")

        station_info = STATION_INFO.get(row["station"].upper())
        network_name = station_info["network_name"] if station_info else row["station"]
        spotcable_code = station_info["spotcable_code"] if station_info else None

        network = _sub(detail_line, "network")
        _sub(network, "name", network_name)
        net_id = _sub(network, "ID")
        _sub(net_id, "code", row["station"], codeOwner="Spotcable", codeDescription=spotcable_code)
        net_id = _sub(network, "ID")
        _sub(net_id, "code", network_name, codeOwner="Strata", codeDescription="Station")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", "TV", codeOwner="Strata", codeDescription="Band")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", start.isoformat(), codeOwner="Strata", codeDescription="Start Date")
        net_id = _sub(network, "ID")
        _sub(net_id, "code", end.isoformat(), codeOwner="Strata", codeDescription="End Date")

        _sub(detail_line, "spotCost", f"{row['rate']:.2f}", currency="USD")

        demo_value = _sub(detail_line, "demoValue", demoRank="1")
        _sub(demo_value, "value", f"{float(row['rating']):.1f}", type="Ratings")
        impressions = round(row["rating"] * population / 100) if population is not None else 0
        _sub(demo_value, "value", impressions, type="Impressions")

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

    return warnings
