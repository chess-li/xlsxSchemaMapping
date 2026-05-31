import re
from typing import Any


SUFFIX_MARKERS = ("必填", "可选", "选填", "required", "optional")
PUNCTUATION_PATTERN = re.compile(r"[\s_\-—–*()（）\[\]【】{}<>《》:：,，.。/\\|]+")


def normalize_field_name(text: Any) -> str:
    value = "" if text is None else str(text)
    value = value.strip().lower()
    for marker in SUFFIX_MARKERS:
        value = value.replace(marker, "")
    return PUNCTUATION_PATTERN.sub("", value)
