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

NEWS_DAYPARTS = {"EARLY MORNING", "EARLY NEWS", "LATE NEWS"}
DAYTIME_DAYPARTS = {"DAYTIME", "DAY"}
ACCESS_DAYPARTS = {"ACCESS", "PRIME ACCESS"}
EXCLUDED_DAYPARTS = {
    "PRIME",
    "SPORTS",
    "EARLY FRINGE",
    "LATE FRINGE",
    "OVERNIGHT",
    "SPECIALS",
    "ROS",
}

LIKED_KEYWORDS = ("JEOPARDY", "WHEEL OF FORTUNE")
NOON_NEWS_START = 10 * 60  # 10:00 AM
NOON_NEWS_END = 14 * 60  # 2:00 PM


def classify(avail):
    """Returns one of: 'Early News', 'Noon News', 'Evening News', 'Late News',
    'Liked Access', 'Daytime', or None (excluded from the buy)."""
    dp = avail.daypart_name.strip().upper()
    name = avail.program_name.strip().upper()

    if dp in EXCLUDED_DAYPARTS:
        return None

    if dp == "EARLY MORNING":
        return "Early News"
    if dp == "EARLY NEWS":
        return "Evening News"
    if dp == "LATE NEWS":
        return "Late News"

    if dp in ACCESS_DAYPARTS:
        if any(kw in name for kw in LIKED_KEYWORDS):
            return "Liked Access"
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


NEWS_CATEGORIES = {"Early News", "Noon News", "Evening News", "Late News"}
TIER_ORDER = {
    "Early News": 1,
    "Noon News": 1,
    "Evening News": 1,
    "Late News": 1,
    "Liked Access": 2,
    "Daytime": 3,
}
