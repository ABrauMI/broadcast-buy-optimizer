"""Shared row-grouping logic used by both the Excel flowchart and the
Strata order export, so they can never disagree about what was bought."""

from collections import defaultdict

from .classify import daypart_code


def group_avails_into_rows(eligible_avails, spots):
    """One row per unique (category, station, program, time, rate, rating),
    covering every eligible avail -- not just the ones actually bought.
    The rate/rating are part of the identity because some rate cards quote
    the same program/time slot at different prices across the quarter (a
    rate escalation, or day-pattern variants that happen to share identical
    clock times) -- folding those into a single row would let the row's one
    displayed rate silently mismatch what specific days were actually
    bought at. Days the program doesn't air are omitted; days it airs but
    wasn't bought get an explicit 0 so the buy can be edited in place."""
    identities = {}
    for a in eligible_avails:
        key = (a._category, a.station, a.program_name, a.start_min, a.end_min, a.rate, a.rating)
        info = identities.setdefault(
            key, {"days": set(), "daypart_name": a.daypart_name, "spot_length": a.spot_length}
        )
        info["days"].update(a.days)

    bought_counts = defaultdict(lambda: defaultdict(int))
    for s in spots:
        key = (s["category"], s["station"], s["program_name"], s["start_min"], s["end_min"], s["rate"], s["rating"])
        bought_counts[key][s["day"]] += 1

    rows = []
    for (category, station, program, start_min, end_min, rate, rating), info in identities.items():
        counts = bought_counts.get((category, station, program, start_min, end_min, rate, rating), {})
        day_counts = {d: counts.get(d, 0) for d in info["days"]}
        rows.append(
            {
                "category": category,  # internal buy tier -- drives the fill color
                "daypart": daypart_code(info["daypart_name"]),  # normalized to the 2-letter code, for the Excel sheet
                "daypart_name": info["daypart_name"],  # the rate card's own label, for exports that need the full name
                "station": station,
                "program": program,
                "start_min": start_min,
                "end_min": end_min,
                "day_counts": day_counts,
                "rate": rate,
                "rating": rating,
                "spot_length": info["spot_length"],
            }
        )
    rows.sort(key=lambda r: (r["station"], r["start_min"], r["program"]))
    return rows
