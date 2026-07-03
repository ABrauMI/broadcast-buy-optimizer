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
    "SPORTS",
    "EARLY FRINGE",
    "LATE FRINGE",
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

    if dp == "PRIME":
        # Primetime is excluded from the buy by default, except the specific
        # programs called out as exceptions: Wheel/Jeopardy, and prime news
        # programming like 60 Minutes. The rest of Prime still shows up in
        # the flowchart (for visibility) under its own category, which the
        # builder never buys into -- it'll always carry zeros.
        if any(kw in name for kw in LIKED_KEYWORDS):
            return "Liked Access"
        if any(kw in name for kw in PRIME_NEWS_KEYWORDS):
            return "Prime News"
        return "Prime"

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
