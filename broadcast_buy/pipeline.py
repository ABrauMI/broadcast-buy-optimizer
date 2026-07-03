"""Shared parse -> build -> export pipeline, used by both the CLI and the
Slack app so neither has to duplicate the wiring."""

from .builder import build_sample_buy
from .excel_export import write_workbook
from .parser import parse_rate_card


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
):
    """Parses every rate card, builds the sample buy, and writes the workbook
    to output_path. Returns (result, log_lines) -- log_lines mirrors what the
    CLI prints, for callers (like the Slack app) that want to surface it."""
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

    return result, log_lines
