from excel_matcher.config import Settings
from excel_matcher.matcher.fusion_matcher import FusionMatcher
from excel_matcher.models import ColumnProfile, FieldMappingResult, StandardField


def match_profiles(
    profiles: list[ColumnProfile],
    standard_fields: list[StandardField],
    enable_embedding: bool = False,
    enable_llm: bool = False,
    settings: Settings | None = None,
    embedding_matcher=None,
    llm_matcher=None,
) -> list[FieldMappingResult]:
    matcher = FusionMatcher(
        standard_fields=standard_fields,
        enable_embedding=enable_embedding,
        enable_llm=enable_llm,
        settings=settings,
        embedding_matcher=embedding_matcher,
        llm_matcher=llm_matcher,
    )
    return [matcher.match(profile) for profile in profiles]
