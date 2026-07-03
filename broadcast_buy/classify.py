"""Maps a rate-card Avail onto Andrew's buy categories.

Buy priority tiers (highest first):
  1. NEWS   -- early (morning), noon, evening, late newscasts. Bought first,
              largely price-insensitive except for extreme CPP outliers.
  2. LIKED  -- Wheel of Fortune / Jeopardy specifically, in their normal
              access-hour syndication slot (not primetime specials).
  3. DAYTIME -- remaining daytime programming, filled in efficiency order.

Anything not matched to one of the above (primetime, sports, early/late
fringe, overnight, specials, digital/streaming "ROS" inventory) is EXCLUDED.
"""

EARLY_MORNING_DAYPARTS = {"EARLY MORNING", "EM"}
EARLY_NEWS_DAYPARTS = {"EARLY NEWS", "EN"}
LATE_NEWS_DAYPARTS = {"LATE NEWS", "LN"}
DAYTIME_DAYPARTS = {"DAYTIME", "DAY", "DY"}
ACCESS_DAYPARTS = {"ACCESS", "PRIME ACCESS", "PA"}
PRIME_DAYPARTS = {"PRIME", "PR"}
EXCLUDED_DAYPARTS = {
    "SPORTS",
    "SP",
    "EARLY FRINGE",
    "EF",
    "LATE FRINGE",
    "LF",
    "OVERNIGHT",
    "SPECIALS",
    "ROS",
}

LIKED_KEYWORDS = ("JEOPARDY", "WHEEL OF FORTUNE")
PRIME_NEWS_KEYWORDS = ("60 MINUTES",)
NOON_NEWS_START = 10 * 60  # 10:00 AM
NOON_NEWS_END = 14 * 60  # 2:00 PM


def classify(avail):
    """Returns one of: 'Early News', 'Noon News', 'Evening News', 'Late News',
    'Prime News', 'Liked Access', 'Daytime', 'Prime' (shown but never bought),
    or None (excluded from the buy and from the flowchart entirely)."""
    dp = avail.daypart_name.strip().upper()
    name = avail.program_name.strip().upper()

    # Wheel/Jeopardy and prime news exceptions are checked before any
    # daypart-based exclusion, since different stations file the exact same
    # syndicated show under different dayparts (Access, Prime, or even Early
    # Fringe depending on the market's clearance) -- the show matters more
    # than where a given station happens to schedule it.
    if any(kw in name for kw in LIKED_KEYWORDS):
        return "Liked Access"
    if any(kw in name for kw in PRIME_NEWS_KEYWORDS):
        return "Prime News"

    if dp in PRIME_DAYPARTS:
        # Everything else in primetime is excluded from the buy, but still
        # shown in the flowchart (for visibility) under its own category --
        # the builder never schedules it, so it always carries zeros.
        return "Prime"

    if dp in EXCLUDED_DAYPARTS:
        return None

    if dp in EARLY_MORNING_DAYPARTS:
        return "Early News"
    if dp in EARLY_NEWS_DAYPARTS:
        return "Evening News"
    if dp in LATE_NEWS_DAYPARTS:
        return "Late News"

    if dp in ACCESS_DAYPARTS:
        return None  # other access-hour programming isn't called out by the guidelines

    if dp in DAYTIME_DAYPARTS:
        if "NEWS" in name and NOON_NEWS_START <= avail.start_min < NOON_NEWS_END:
            return "Noon News"
        if "NEWS" in name:
            # a news-branded daytime program outside the noon window (rare) --
            # still news, fold into Evening/Early based on which side it's closer to
            return "Noon News"
        return "Daytime"

    return None
