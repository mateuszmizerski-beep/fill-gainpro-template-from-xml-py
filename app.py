import importlib
from pathlib import Path
import re
import tempfile

import streamlit as st

import fill_gainpro_template_from_xml as extractor


# Streamlit may rerun this file without refreshing an already-imported helper module.
extractor = importlib.reload(extractor)
XmlFinancials = extractor.XmlFinancials
build_fill_jobs = extractor.build_fill_jobs
fill_template = extractor.main


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "assets" / "gainpro_template.xlsx"
MAX_FILES = 5
MAX_FILE_SIZE_MB = 25
LEGAL_FORM_SUFFIX = re.compile(
    r"\s+(?:"
    r"spółka\b.*"
    r"|sp\.?\s*z\.?\s*o\.?\s*o\.?.*"
    r"|sp\.?\s*k\.?.*"
    r"|sp\.?\s*j\.?.*"
    r"|s\.?\s*k\.?\s*a\.?.*"
    r"|s\.?\s*a\.?.*"
    r")$",
    re.IGNORECASE,
)


def output_filename(xml_path: Path) -> str:
    company = XmlFinancials(xml_path).company or "Filled"
    safe_company = re.sub(r'[<>:"/\\|?*]+', "", company).strip()
    safe_company = re.sub(r"\s+", " ", safe_company)
    safe_company = LEGAL_FORM_SUFFIX.sub("", safe_company).strip()
    if not safe_company:
        safe_company = "Filled"
    elif safe_company.isupper():
        safe_company = safe_company.title()
    safe_company = safe_company[:120].strip() or "Filled"
    return f"{safe_company} Financials.xlsx"


st.set_page_config(
    page_title="Gain Financials XML Extractor",
)

st.title("Gain Financials XML Extractor")
st.write(
    "Upload up to 5 XML files for annual reports from KRS. "
    "Order and file names do not matter."
)
st.warning(
    'Use normal XML files only. Files ending in ".xml.xades" will not work. '
    'If KRS provides an .xades file, select "Pobierz treść dokumentu" '
    "to download the normal XML file."
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
                years_to_fill = str(len(xml_paths) + 1)

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
                planned_jobs = build_fill_jobs(
                    [XmlFinancials(path) for path in xml_paths],
                    target_year=None,
                    fill_comparative=should_fill_comparative,
                    years=int(years_to_fill),
                )
                annualised_jobs = [
                    job
                    for job in planned_jobs
                    if job.period_xml_data and job.period_xml_data.requires_annualisation
                ]
                fill_template(args)
                download_filename = output_filename(xml_paths[0])
            except SystemExit as exc:
                st.error(f"The workbook could not be created: {exc}")
                st.stop()
            except Exception as exc:
                st.error(f"Unexpected processing error: {exc}")
                st.stop()

            result = output_path.read_bytes()

    if annualised_jobs:
        periods = ", ".join(
            (
                f"FY{job.year} "
                f"({job.period_xml_data.period_start_date:%d-%m-%Y} to "
                f"{job.period_xml_data.period_end_date:%d-%m-%Y})"
            )
            for job in annualised_jobs
        )
        st.warning(
            f"Annualisation detected for {periods}. The workbook has been generated "
            "with annualised flow figures. Please check the dates, annualisation "
            "factor, P&L and cash flow values particularly carefully. "
            "Important: a broken fiscal year can only be detected when the XML for "
            "that fiscal year is uploaded. A later full-year XML containing comparative "
            "figures does not provide the prior report's start and end dates."
        )

    st.success("Your filled Excel workbook is ready.")

    st.download_button(
        label="Download Filled Excel Workbook",
        data=result,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("made by Bronek xoxo")
st.caption("Questions? Message me on Slack.")
