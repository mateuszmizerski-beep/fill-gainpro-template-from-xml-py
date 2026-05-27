from pathlib import Path
import tempfile

import streamlit as st

from fill_gainpro_template_from_xml import main as fill_template


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "assets" / "gainpro_template.xlsx"
MAX_FILES = 5
MAX_FILE_SIZE_MB = 25


st.set_page_config(
    page_title="Gain Financials XML Extractor",
)

st.title("Gain Financials XML Extractor")
st.write(
    "Upload up to 5 XML files for annual reports from KRS. "
    "Order and file names do not matter."
)

with st.form("xml_upload_form"):
    xml_files = st.file_uploader(
        "XML annual report files",
        type=["xml"],
        accept_multiple_files=True,
    )

    with st.expander("Advanced settings"):
        skip_single_xml_comparative = st.checkbox(
            "Do not fill the previous-year comparative figures when uploading one XML",
            value=False,
        )

    submitted = st.form_submit_button("Generate Excel")

if submitted:
    if not xml_files:
        st.warning("Please upload at least one XML file.")
        st.stop()

    if len(xml_files) > MAX_FILES:
        st.error(f"Please upload no more than {MAX_FILES} XML files.")
        st.stop()

    oversized_files = [
        uploaded.name
        for uploaded in xml_files
        if uploaded.size > MAX_FILE_SIZE_MB * 1024 * 1024
    ]
    if oversized_files:
        st.error(
            f"Each XML must be smaller than {MAX_FILE_SIZE_MB} MB. "
            f"Too large: {', '.join(oversized_files)}"
        )
        st.stop()

    if not TEMPLATE_PATH.exists():
        st.error("The embedded Excel template is missing from the app.")
        st.stop()

    with st.spinner("Creating filled Excel workbook..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)
            xml_paths = []

            for index, uploaded_file in enumerate(xml_files, start=1):
                xml_path = tmpdir_path / f"input_{index}.xml"
                xml_path.write_bytes(uploaded_file.getvalue())
                xml_paths.append(xml_path)

            output_path = tmpdir_path / "Filled_Financials.xlsx"
            should_fill_comparative = (
                len(xml_paths) == 1 and not skip_single_xml_comparative
            )

            if len(xml_paths) == 1:
                years_to_fill = "2" if should_fill_comparative else "1"
            else:
                years_to_fill = "5"

            args = [
                "--xml",
                *[str(path) for path in xml_paths],
                "--template",
                str(TEMPLATE_PATH),
                "--output",
                str(output_path),
                "--years",
                years_to_fill,
            ]

            if should_fill_comparative:
                args.append("--fill-comparative")

            try:
                fill_template(args)
            except SystemExit as exc:
                st.error(f"The workbook could not be created: {exc}")
                st.stop()
            except Exception as exc:
                st.error(f"Unexpected processing error: {exc}")
                st.stop()

            result = output_path.read_bytes()

    st.success("Your filled Excel workbook is ready.")

    st.download_button(
        label="Download Filled Excel Workbook",
        data=result,
        file_name="Filled_Financials.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption(
    "The workbook may recalculate formulas when it is first opened in Microsoft Excel."
)
