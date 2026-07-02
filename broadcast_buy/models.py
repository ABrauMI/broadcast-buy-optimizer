from dataclasses import dataclass, field

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


@dataclass
class Avail:
    """One line item (program/time slot) from a station rate card."""

    station: str
    daypart_name: str
    program_name: str
    start_min: int  # minutes from midnight; may exceed 1440 for post-midnight slots
    end_min: int
    days: list  # subset of DAYS this avail airs on
    rate: float
    rating: float | None  # target-demo rating, None if station doesn't sell it
    spot_length: str = ":30"

    @property
    def cpp(self):
        if not self.rating or self.rating <= 0:
            return None
        return self.rate / self.rating


@dataclass
class ScheduledSpot:
    station: str
    category: str
    program_name: str
    day: str
    start_min: int
    end_min: int
    rate: float
    rating: float

    @property
    def cpp(self):
        return self.rate / self.rating if self.rating else None
