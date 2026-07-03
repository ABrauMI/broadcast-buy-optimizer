"""Shared parse -> build -> export pipeline, used by both the CLI and the
Slack app so neither has to duplicate the wiring."""

import os

from .builder import build_sample_buy
from .excel_export import write_workbook
from .grouping import group_avails_into_rows
from .parser import parse_rate_card
from .strata_export import write_strata_order
from .workbook_import import read_sample_buy_rows


def parse_clock(text):
    """Accepts 'HH:MM' or a bare hour like '7' / '23'."""
    if ":" in text:
        h, m = text.split(":")
        return int(h) * 60 + int(m)
    return int(text) * 60


def run_pipeline(
    xml_paths,
    output_path,
    target_grps=750.0,
    target_demo_group="Adults",
    target_demo_age=35,
    earliest_time="7:00",
    latest_time="23:00",
    market_name=None,
    flight_start=None,
    flight_end=None,
    campaign_name=None,
    strata_output_path=None,
):
    """Parses every rate card, builds the sample buy, and writes the workbook
    to output_path. Returns (result, log_lines, strata_path) -- log_lines
    mirrors what the CLI prints, for callers (like the Slack app) that want
    to surface it. strata_path is None unless market_name/flight_start/
    flight_end are all supplied, in which case a Strata-importable .sbx
    order is also written (defaulting next to output_path) and its path
    returned."""
    all_avails = []
    log_lines = []
    for path in xml_paths:
        station, avails, warnings = parse_rate_card(
            path, target_group=target_demo_group, target_age_from=target_demo_age
        )
        for w in warnings:
            log_lines.append(f"WARNING: {w}")
        log_lines.append(f"Parsed {station}: {len(avails)} avail rows from {path}")
        all_avails.extend(avails)

    result = build_sample_buy(
        all_avails,
        target_grps=target_grps,
        earliest_min=parse_clock(earliest_time),
        latest_min=parse_clock(latest_time),
    )

    demo_label = f"{target_demo_group} {target_demo_age}+"
    write_workbook(result, output_path, target_demo_label=demo_label)

    strata_path = None
    if market_name and flight_start and flight_end:
        strata_path = strata_output_path or os.path.splitext(output_path)[0] + ".sbx"
        rows = group_avails_into_rows(result.eligible_avails, result.spots)
        strata_warnings = write_strata_order(
            rows,
            strata_path,
            market_name=market_name,
            flight_start=flight_start,
            flight_end=flight_end,
            target_demo_group=target_demo_group,
            target_demo_age=target_demo_age,
            campaign_name=campaign_name,
        )
        for w in strata_warnings:
            log_lines.append(f"STRATA WARNING: {w}")

    return result, log_lines, strata_path


def build_strata_order_from_workbook(
    xlsx_path,
    output_path,
    *,
    market_name,
    flight_start,
    flight_end,
    target_demo_group="Adults",
    target_demo_age=35,
    campaign_name=None,
):
    """Reads a Sample Buy workbook back in -- including one a buyer opened
    and edited by hand (changed day quantities, rates, or added/removed
    rows) -- and writes a Strata .sbx order from whatever it currently
    contains. This is the edit-then-export path, as opposed to
    run_pipeline's rate-cards-to-both-outputs path. Returns log_lines for
    the caller to surface, same convention as run_pipeline."""
    log_lines = []
    rows = read_sample_buy_rows(xlsx_path)
    log_lines.append(f"Read {len(rows)} buy row(s) from {xlsx_path}")
    warnings = write_strata_order(
        rows,
        output_path,
        market_name=market_name,
        flight_start=flight_start,
        flight_end=flight_end,
        target_demo_group=target_demo_group,
        target_demo_age=target_demo_age,
        campaign_name=campaign_name,
    )
    for w in warnings:
        log_lines.append(f"STRATA WARNING: {w}")
    return log_lines
