"""Parser for AAAA/TVB SpotTVCableProposal rate-card XML (the format Strata imports)."""

import xml.etree.ElementTree as ET

from .models import Avail, DAYS

NS_MAIN = "http://www.AAAA.org/schemas/spotTVCableProposal"
NS_TVB = "http://www.AAAA.org/schemas/spotTV"
NS_TVBTP = "http://www.AAAA.org/schemas/TVBGeneralTypes"

DAY_TAGS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def q(tag):
    return f"{{{NS_MAIN}}}{tag}"


def qtvb(tag):
    return f"{{{NS_TVB}}}{tag}"


def qtp(tag):
    return f"{{{NS_TVBTP}}}{tag}"


def parse_time(text):
    """'HH:MM' -> minutes from midnight. Hours may run past 24 for post-midnight slots
    (broadcast-day convention, e.g. 28:00 = 4:00 AM the next calendar day)."""
    h, m = text.split(":")
    return int(h) * 60 + int(m)


def find_target_demo_id(demo_categories_el, group, age_from, demo_type="Rating"):
    for demo_cat in demo_categories_el.findall(q("DemoCategory")):
        d_type = demo_cat.findtext(qtvb("DemoType"))
        d_group = demo_cat.findtext(qtvb("Group"))
        d_age_from = demo_cat.findtext(qtvb("AgeFrom"))
        if (
            d_type == demo_type
            and d_group == group
            and d_age_from is not None
            and int(d_age_from) == age_from
        ):
            return demo_cat.get("DemoId")
    return None


def parse_rate_card(path, station_override=None, target_group="Adults", target_age_from=35):
    """Returns (station_call_letters, list[Avail], warnings)."""
    tree = ET.parse(path)
    root = tree.getroot()
    warnings = []

    proposal = root.find(q("Proposal"))
    outlet = proposal.find(q("Outlets")).find(q("TelevisionStation"))
    station = station_override or outlet.get("callLetters")

    avail_list = proposal.find(q("AvailList"))
    demo_categories_el = avail_list.find(q("DemoCategories"))
    target_demo_id = find_target_demo_id(demo_categories_el, target_group, target_age_from)
    if target_demo_id is None:
        warnings.append(
            f"{station}: no {target_group} {target_age_from}+ Rating demo found in rate card; "
            "ratings will be blank for all avails."
        )

    avails = []
    for line in avail_list.findall(q("AvailLineWithDetailedPeriods")):
        daypart_name = (line.findtext(q("DaypartName")) or "").strip()
        program_name = (line.findtext(q("AvailName")) or "").strip()
        spot_length = line.findtext(q("SpotLength")) or ":30"

        periods = line.find(q("Periods"))
        detailed_period = periods.find(q("DetailedPeriod")) if periods is not None else None
        if detailed_period is None:
            continue
        rate_text = detailed_period.findtext(q("Rate"))
        rate = float(rate_text) if rate_text else 0.0

        rating = None
        demo_values = detailed_period.find(q("DemoValues"))
        if demo_values is not None and target_demo_id is not None:
            for dv in demo_values.findall(q("DemoValue")):
                if dv.get("demoRef") == target_demo_id:
                    rating = float(dv.text)
                    break

        day_times = line.find(q("DayTimes"))
        for day_time in day_times.findall(q("DayTime")):
            start_text = day_time.findtext(q("StartTime"))
            end_text = day_time.findtext(q("EndTime"))
            try:
                start_min = parse_time(start_text)
                end_min = parse_time(end_text)
            except (ValueError, AttributeError):
                warnings.append(
                    f"{station}: could not parse time for '{program_name}' "
                    f"({start_text}-{end_text}), skipped."
                )
                continue

            days_el = day_time.find(q("Days"))
            active_days = []
            for day_name, day_tag in zip(DAYS, DAY_TAGS):
                flag = days_el.findtext(qtp(day_tag))
                if flag == "Y":
                    active_days.append(day_name)

            if not active_days:
                continue

            avails.append(
                Avail(
                    station=station,
                    daypart_name=daypart_name,
                    program_name=program_name,
                    start_min=start_min,
                    end_min=end_min,
                    days=active_days,
                    rate=rate,
                    rating=rating,
                    spot_length=spot_length,
                )
            )

    return station, avails, warnings
