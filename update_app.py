import importlib
from pathlib import Path
import re
import tempfile

import streamlit as st

import fill_gainpro_template_from_xml as extractor
import update_gainpro_template_from_existing_and_xml as updater


# Streamlit can rerun this file while helper modules remain cached.
extractor = importlib.reload(extractor)
updater = importlib.reload(updater)

XmlFinancials = extractor.XmlFinancials
build_fill_jobs = extractor.build_fill_jobs
update_workbook = updater.update_workbook


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "assets" / "gainpro_template.xlsx"
MAX_XML_FILES = 5
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


def clean_company_name(company: str) -> str:
    safe_company = re.sub(r'[<>:"/\\|?*]+', "", company).strip()
    safe_company = re.sub(r"\s+", " ", safe_company)
    safe_company = LEGAL_FORM_SUFFIX.sub("", safe_company).strip()
    if safe_company.isupper():
        safe_company = safe_company.title()
    return safe_company[:120].strip()


def output_filename(xml_paths: list[Path], existing_file_name: str) -> str:
    company = ""
    if xml_paths:
        try:
            company = XmlFinancials(xml_paths[-1]).company or ""
        except Exception:
            company = ""

    safe_company = clean_company_name(company)
    if not safe_company:
        safe_company = clean_company_name(Path(existing_file_name).stem)
    if not safe_company:
        safe_company = "Updated"
    return f"{safe_company} Financials Updated.xlsx"


def safe_uploaded_filename(file_name: str, fallback: str) -> str:
    safe_name = Path(file_name).name
    safe_name = re.sub(r'[<>:"/\\|?*]+', "_", safe_name).strip()
    return safe_name or fallback


def planned_update_jobs(xml_paths: list[Path], existing_years: list[int]):
    xml_files = [XmlFinancials(path) for path in xml_paths]
    newest_xml_year = max(xml_data.year for xml_data in xml_files)
    latest_existing_year = max(existing_years) if existing_years else newest_xml_year - 1
    fill_year_count = max(2, newest_xml_year - latest_existing_year)
    return build_fill_jobs(
        xml_files,
        target_year=None,
        fill_comparative=True,
        years=fill_year_count,
    )


st.set_page_config(
    page_title="Gain Financials XML Updater",
)

st.title("Gain Financials XML Updater")
st.write(
    "Upload an existing filled Gain financials Excel and one or more new KRS "
    "annual report XML files. The app will move the existing data into a fresh "
    "template and add or refresh the newest financial years."
)
st.info(
    "FTEs and Segmentations will never be filled out because these are not "
    "included in XMLs. These fields still need to be filled out manually. "
    "Existing Segmentations from the uploaded Excel are copied over."
)
st.warning(
    'Use normal XML files only. Files ending in ".xml.xades" will not work. '
    'If KRS provides an .xades file, select "Pobierz treść dokumentu" '
    "to download the normal XML file."
)

with st.form("update_upload_form"):
    existing_excel = st.file_uploader(
        "Existing filled Gain financials Excel",
        type=["xlsx"],
        accept_multiple_files=False,
    )
    xml_files = st.file_uploader(
        "New XML annual report files",
        type=["xml"],
        accept_multiple_files=True,
    )

    with st.expander("Advanced settings"):
        allow_company_mismatch = st.checkbox(
            "Allow XML company name to differ from the Excel file name",
            value=False,
        )

    submitted = st.form_submit_button("Generate Updated Excel")

if submitted:
    if existing_excel is None:
        st.warning("Please upload the existing filled Excel workbook.")
        st.stop()

    if not xml_files:
        st.warning("Please upload at least one XML file.")
        st.stop()

    if len(xml_files) > MAX_XML_FILES:
        st.error(f"Please upload no more than {MAX_XML_FILES} XML files.")
        st.stop()

    oversized_files = [
        uploaded.name
        for uploaded in [existing_excel, *xml_files]
        if uploaded.size > MAX_FILE_SIZE_MB * 1024 * 1024
    ]
    if oversized_files:
        st.error(
            f"Each uploaded file must be smaller than {MAX_FILE_SIZE_MB} MB. "
            f"Too large: {', '.join(oversized_files)}"
        )
        st.stop()

    if not TEMPLATE_PATH.exists():
        st.error("The embedded Excel template is missing from the app.")
        st.stop()

    with st.spinner("Creating updated Excel workbook..."):
        with tempfile.TemporaryDirectory() as tmpdir:
            tmpdir_path = Path(tmpdir)

            existing_path = tmpdir_path / safe_uploaded_filename(
                existing_excel.name,
                "existing.xlsx",
            )
            existing_path.write_bytes(existing_excel.getvalue())

            xml_paths: list[Path] = []
            for index, uploaded_file in enumerate(xml_files, start=1):
                xml_path = tmpdir_path / f"input_{index}.xml"
                xml_path.write_bytes(uploaded_file.getvalue())
                xml_paths.append(xml_path)

            output_path = tmpdir_path / "Updated_Financials.xlsx"

            try:
                existing_wb = updater.load_workbook(existing_path, read_only=True, data_only=False)
                existing_ws = existing_wb[updater.FINANCIALS_SHEET]
                existing_years = updater.active_years(existing_ws)
                planned_jobs = planned_update_jobs(xml_paths, existing_years)
                annualised_jobs = [
                    job
                    for job in planned_jobs
                    if job.period_xml_data and job.period_xml_data.requires_annualisation
                ]

                update_workbook(
                    existing_path,
                    TEMPLATE_PATH,
                    xml_paths,
                    output_path,
                    allow_company_mismatch=allow_company_mismatch,
                )
                download_filename = output_filename(xml_paths, existing_excel.name)
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

    st.success("Your updated Excel workbook is ready.")

    st.download_button(
        label="Download Updated Excel Workbook",
        data=result,
        file_name=download_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.caption("made by Bronek xoxo")
st.caption("Questions? Message me on Slack.")
