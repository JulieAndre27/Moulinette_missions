"""other small functions"""


def split_strip_lower(s: str, splitter: str = ",") -> list[str]:
    """lower then split <s> with <splitter>, and strip each member"""
    return [i.strip() for i in s.lower().split(splitter)]
