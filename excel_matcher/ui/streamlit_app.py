import tempfile
from pathlib import Path

import streamlit as st

from excel_matcher.config import Settings
from excel_matcher.services.matching_service import match_profiles
from excel_matcher.services.template_service import build_template_signature
from excel_matcher.services.workbook_service import analyze_workbook
from excel_matcher.storage.default_fields import get_default_fields
from excel_matcher.storage.sqlite_store import SQLiteStore


def main() -> None:
    st.set_page_config(page_title="Excel Matcher MVP", layout="wide")
    st.title("Excel Matcher MVP")

    uploaded_file = st.file_uploader("Upload .xlsx", type=["xlsx"])
    if uploaded_file is None:
        return

    workbook_path = _save_upload(uploaded_file)
    analysis = analyze_workbook(workbook_path)
    standard_fields = get_default_fields()
    field_options = [""] + [field.key for field in standard_fields]

    st.subheader("Workbook")
    st.write(
        {
            "path": analysis["workbook"].path,
            "sheets": [sheet.name for sheet in analysis["workbook"].sheets],
        }
    )

    confirmed_mappings = {}
    for sheet in analysis["workbook"].sheets:
        st.subheader(sheet.name)
        header = analysis["headers"][sheet.name]
        profiles = analysis["profiles"][sheet.name]
        st.write(
            {
                "title_row": header.title_row,
                "header_row": header.header_row,
                "data_start_row": header.data_start_row,
                "confidence": header.confidence,
                "evidence": header.evidence,
            }
        )
        st.dataframe([profile.model_dump() for profile in profiles], use_container_width=True)

        mappings = match_profiles(
            profiles,
            standard_fields,
            enable_embedding=False,
            enable_llm=False,
        )
        for mapping in mappings:
            with st.expander(f"{mapping.excel_field} -> {mapping.target_field or ''}"):
                selected = st.selectbox(
                    "Confirmed target",
                    field_options,
                    index=_selected_index(field_options, mapping.target_field),
                    key=f"{sheet.name}:{mapping.excel_field}",
                )
                if selected:
                    confirmed_mappings[mapping.excel_field] = selected
                st.write(
                    {
                        "confidence": mapping.confidence,
                        "needs_review": mapping.needs_review,
                        "candidates": [candidate.model_dump() for candidate in mapping.candidates],
                    }
                )
                st.dataframe(
                    [
                        {
                            "stage": stage.stage,
                            "status": stage.status.value,
                            "reason": stage.reason,
                            "candidates": [candidate.model_dump() for candidate in stage.candidates],
                        }
                        for stage in mapping.stage_results
                    ],
                    use_container_width=True,
                )

        if st.button(f"Save template for {sheet.name}"):
            signature = build_template_signature(sheet.name, profiles)
            store = SQLiteStore(Settings().sqlite_path)
            store.initialize()
            store.save_template(signature.signature, sheet.name, confirmed_mappings)
            st.success("Template saved")


def _save_upload(uploaded_file) -> Path:
    suffix = Path(uploaded_file.name).suffix or ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temporary_file:
        temporary_file.write(uploaded_file.getbuffer())
        return Path(temporary_file.name)


def _selected_index(options: list[str], target_field: str | None) -> int:
    if target_field in options:
        return options.index(target_field)
    return 0


if __name__ == "__main__":
    main()
