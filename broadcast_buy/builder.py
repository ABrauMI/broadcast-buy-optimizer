"""Builds a sample weekly buy targeting a market GRP goal, following the
priority rules: news first (price-insensitive, but the craziest CPP outliers
get deprioritized), then Wheel/Jeopardy, then the most efficient/best-rated
remaining daytime, all while respecting a 30-minute minimum spacing between
any two spots on the same station on the same day.

Long-format programs can carry more than one spot per airing -- e.g. a
2-hour block can fit spots every 30 minutes (4 total), an hour-long one can
fit 2 -- but only when the avail is cost-efficient (at/below the tier's
median CPP for news, or anywhere in the cost-sorted daytime tier). Cost
exceptions bought regardless of price (Prime News, Liked Access) stay
capped at one spot per airing.
"""

import statistics
from collections import defaultdict
from dataclasses import dataclass, field

from .classify import classify

MIN_SPOT_GAP_MIN = 30
OUTLIER_CPP_MULTIPLIER = 3.0
DEFAULT_EARLIEST_MIN = 7 * 60  # 7:00 AM
DEFAULT_LATEST_MIN = 23 * 60  # 11:00 PM
NEWS_CORE_CATEGORIES = {"Early News", "Noon News", "Evening News", "Late News"}


@dataclass
class BuyResult:
    spots: list = field(default_factory=list)  # list[ScheduledSpot-like dict]
    eligible_avails: list = field(default_factory=list)  # every avail considered, bought or not
    target_grps: float = 0.0
    achieved_grps: float = 0.0
    total_cost: float = 0.0
    warnings: list = field(default_factory=list)
    outlier_avails: list = field(default_factory=list)


def _fits(scheduled_starts, day, start, min_gap=MIN_SPOT_GAP_MIN):
    """Spacing is judged by spot start times, not avail duration -- this is
    what lets several spots share one long program (each 30 min apart)
    while still blocking spots in two different programs from landing
    closer than the minimum gap."""
    return all(abs(start - existing) >= min_gap for existing in scheduled_starts.get(day, []))


def _max_spots_for_avail(avail, efficient):
    if not efficient:
        return 1
    duration = avail.end_min - avail.start_min
    return max(1, duration // MIN_SPOT_GAP_MIN)


def _cpp(avail):
    return avail.rate / avail.rating if avail.rating else None


def _within_time_window(avail, earliest_min, latest_min):
    """True if the avail's whole time window falls within [earliest_min, latest_min)
    clock time. Avails that wrap past midnight are treated as violating the
    window (they necessarily extend into the excluded overnight hours)."""
    start = avail.start_min % 1440
    end = avail.end_min % 1440
    if end <= start:
        return False
    return start >= earliest_min and end <= latest_min


def _round_robin(groups):
    """groups: dict[key -> list[avail]]. Yields avails interleaved one-per-key per round."""
    queues = {k: list(v) for k, v in groups.items() if v}
    while queues:
        for k in list(queues.keys()):
            if queues[k]:
                yield queues[k].pop(0)
            if not queues[k]:
                del queues[k]


def build_sample_buy(
    all_avails,
    target_grps=750.0,
    earliest_min=DEFAULT_EARLIEST_MIN,
    latest_min=DEFAULT_LATEST_MIN,
):
    result = BuyResult(target_grps=target_grps)

    eligible = [a for a in all_avails if a.rate > 0 and a.rating]
    eligible = [a for a in eligible if _within_time_window(a, earliest_min, latest_min)]
    for a in eligible:
        a._category = classify(a)
    eligible = [a for a in eligible if a._category]
    result.eligible_avails = eligible

    # "Core" news (Early/Noon/Evening/Late) is where the CPP-outlier check
    # applies. "Prime News" (e.g. 60 Minutes) is an explicitly-requested
    # exception to the no-primetime rule, so it's bought at full priority
    # like the rest of news, without being deprioritized for cost.
    news_core = [a for a in eligible if a._category in NEWS_CORE_CATEGORIES]
    prime_news = [a for a in eligible if a._category == "Prime News"]
    liked = [a for a in eligible if a._category == "Liked Access"]
    daytime = [a for a in eligible if a._category == "Daytime"]

    # -- Tier 1: news, non-outliers first, outliers (crazy-high CPP) last --
    news_cpps = [_cpp(a) for a in news_core if _cpp(a) is not None]
    median_news_cpp = statistics.median(news_cpps) if news_cpps else 0
    outlier_threshold = median_news_cpp * OUTLIER_CPP_MULTIPLIER

    news_normal = [a for a in news_core if (_cpp(a) or 0) <= outlier_threshold]
    news_outlier = [a for a in news_core if (_cpp(a) or 0) > outlier_threshold]
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
        _round_robin(
            group_by(news_normal + prime_news, lambda a: (a.station, a._category))
        )
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

    station_day_starts = defaultdict(lambda: defaultdict(list))  # station -> day -> [start_min,...]
    total_grps = 0.0
    total_cost = 0.0
    spots = []

    for avail in ordered_avails:
        if total_grps >= target_grps:
            break

        if avail._category in NEWS_CORE_CATEGORIES:
            cpp_val = _cpp(avail)
            efficient = cpp_val is not None and median_news_cpp > 0 and cpp_val <= median_news_cpp
        elif avail._category == "Daytime":
            efficient = True  # already the cost-sorted, cheapest-first tier
        else:
            efficient = False  # Prime News / Liked Access: bought regardless of cost

        max_spots = _max_spots_for_avail(avail, efficient)

        for day in avail.days:
            if total_grps >= target_grps:
                break
            starts = station_day_starts[avail.station][day]
            placed = 0
            slot_start = avail.start_min
            while placed < max_spots and slot_start < avail.end_min and total_grps < target_grps:
                if _fits(station_day_starts[avail.station], day, slot_start):
                    starts.append(slot_start)
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
                    placed += 1
                slot_start += MIN_SPOT_GAP_MIN

    result.spots = spots
    result.achieved_grps = total_grps
    result.total_cost = total_cost

    if total_grps < target_grps:
        result.warnings.append(
            f"Only reached {total_grps:.1f} of {target_grps:.0f} target weekly GRPs -- "
            "ran out of eligible news/access/daytime inventory under current rules "
            "(30-min spacing cap, 7a-11p daypart window, and daypart exclusions). "
            "Consider loosening spacing, widening the time window, adding Early/Late "
            "Fringe, or adding more stations."
        )

    return result
