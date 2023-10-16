"""other small functions"""


def split_strip_lower(s: str, splitter: str = ",") -> list[str]:
    """lower then split <s> with <splitter>, and strip each member"""
    return [i.strip() for i in s.lower().split(splitter)]


def index_to_column(i: int) -> str:
    """index number to column string, ex 0 -> A, 1 -> B. Doesn't support over Z"""
    return chr(ord("A") + i)
