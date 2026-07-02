"""Builds a sample weekly buy targeting a market GRP goal, following the
priority rules: news first (price-insensitive, but the craziest CPP outliers
get deprioritized), then Wheel/Jeopardy, then the most efficient/best-rated
remaining daytime, all while respecting a 30-minute minimum spacing between
spots on the same station on the same day.
"""

import statistics
from collections import defaultdict
from dataclasses import dataclass, field

from .classify import NEWS_CATEGORIES, TIER_ORDER, classify

MIN_SPOT_GAP_MIN = 30
OUTLIER_CPP_MULTIPLIER = 3.0


@dataclass
class BuyResult:
    spots: list = field(default_factory=list)  # list[ScheduledSpot-like dict]
    target_grps: float = 0.0
    achieved_grps: float = 0.0
    total_cost: float = 0.0
    warnings: list = field(default_factory=list)
    outlier_avails: list = field(default_factory=list)


def _fits(schedule, day, start, end, min_gap=MIN_SPOT_GAP_MIN):
    for e_start, e_end in schedule.get(day, []):
        gap = max(e_start - end, start - e_end)
        if gap < min_gap:
            return False
    return True


def _cpp(avail):
    return avail.rate / avail.rating if avail.rating else None


def _round_robin(groups):
    """groups: dict[key -> list[avail]]. Yields avails interleaved one-per-key per round."""
    queues = {k: list(v) for k, v in groups.items() if v}
    while queues:
        for k in list(queues.keys()):
            if queues[k]:
                yield queues[k].pop(0)
            if not queues[k]:
                del queues[k]


def build_sample_buy(all_avails, target_grps=750.0):
    result = BuyResult(target_grps=target_grps)

    eligible = [a for a in all_avails if a.rate > 0 and a.rating]
    for a in eligible:
        a._category = classify(a)
    eligible = [a for a in eligible if a._category]

    news = [a for a in eligible if a._category in NEWS_CATEGORIES]
    liked = [a for a in eligible if a._category == "Liked Access"]
    daytime = [a for a in eligible if a._category == "Daytime"]

    # -- Tier 1: news, non-outliers first, outliers (crazy-high CPP) last --
    news_cpps = [_cpp(a) for a in news if _cpp(a) is not None]
    median_news_cpp = statistics.median(news_cpps) if news_cpps else 0
    outlier_threshold = median_news_cpp * OUTLIER_CPP_MULTIPLIER

    news_normal = [a for a in news if (_cpp(a) or 0) <= outlier_threshold]
    news_outlier = [a for a in news if (_cpp(a) or 0) > outlier_threshold]
    result.outlier_avails = [
        (a.station, a.program_name, round(_cpp(a), 2)) for a in news_outlier
    ]

    def group_by(avails, key_fn):
        groups = defaultdict(list)
        for a in avails:
            groups[key_fn(a)].append(a)
        for group in groups.values():
            group.sort(key=lambda a: a.start_min)
        return groups

    news_normal_order = list(
        _round_robin(group_by(news_normal, lambda a: (a.station, a._category)))
    )
    news_outlier_order = list(
        _round_robin(group_by(news_outlier, lambda a: (a.station, a._category)))
    )

    # -- Tier 2: Wheel/Jeopardy, spread across stations that carry them --
    liked_order = list(_round_robin(group_by(liked, lambda a: a.station)))

    # -- Tier 3: remaining daytime, cheapest CPP (most efficient) first --
    daytime.sort(key=lambda a: _cpp(a) or float("inf"))
    daytime_order = daytime

    ordered_avails = news_normal_order + news_outlier_order + liked_order + daytime_order

    station_day_schedule = defaultdict(lambda: defaultdict(list))  # station -> day -> [(s,e)]
    total_grps = 0.0
    total_cost = 0.0
    spots = []

    for avail in ordered_avails:
        if total_grps >= target_grps:
            break
        for day in avail.days:
            if total_grps >= target_grps:
                break
            sched = station_day_schedule[avail.station]
            if not _fits(sched, day, avail.start_min, avail.end_min):
                continue
            sched.setdefault(day, []).append((avail.start_min, avail.end_min))
            spots.append(
                {
                    "station": avail.station,
                    "category": avail._category,
                    "program_name": avail.program_name,
                    "day": day,
                    "start_min": avail.start_min,
                    "end_min": avail.end_min,
                    "rate": avail.rate,
                    "rating": avail.rating,
                    "cpp": round(avail.rate / avail.rating, 2),
                }
            )
            total_grps += avail.rating
            total_cost += avail.rate

    result.spots = spots
    result.achieved_grps = total_grps
    result.total_cost = total_cost

    if total_grps < target_grps:
        result.warnings.append(
            f"Only reached {total_grps:.1f} of {target_grps:.0f} target weekly GRPs -- "
            "ran out of eligible news/access/daytime inventory under current rules "
            "(30-min spacing cap and daypart exclusions). Consider loosening spacing, "
            "adding Early/Late Fringe, or adding more stations."
        )

    return result
