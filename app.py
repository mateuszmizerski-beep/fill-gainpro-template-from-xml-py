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
    layout="centered",
)

st.markdown(
    """
    <style>
        .stApp {
            background: var(--background-color);
        }

        .block-container {
            max-width: 760px;
            padding-top: 4.5rem;
            padding-bottom: 3.5rem;
        }

        .extractor-header {
            margin-bottom: 2rem;
        }

        .extractor-label {
            color: var(--primary-color);
            font-size: 0.75rem;
            font-weight: 700;
            letter-spacing: 0.12em;
            margin-bottom: 0.65rem;
            text-transform: uppercase;
        }

        .extractor-title {
            color: var(--text-color) !important;
            font-size: 2.55rem;
            font-weight: 650;
            letter-spacing: -0.045em;
            line-height: 1.14;
            margin: 0 0 0.75rem;
        }

        .extractor-description {
            color: color-mix(in srgb, var(--text-color) 72%, transparent);
            font-size: 1.02rem;
            line-height: 1.55;
            margin: 0;
            max-width: 650px;
        }

        [data-testid="stFileUploader"] {
            margin-top: 0.25rem;
        }

        [data-testid="stFileUploader"] section {
            background: var(--secondary-background-color);
            border: 1px dashed color-mix(in srgb, var(--text-color) 34%, transparent);
            border-radius: 12px;
            padding: 0.55rem;
        }

        [data-testid="stFileUploader"] section:hover {
            border-color: var(--primary-color);
        }

        [data-testid="stFileUploader"] label p {
            color: var(--text-color) !important;
            font-weight: 600;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] span {
            color: var(--text-color) !important;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] small {
            color: color-mix(in srgb, var(--text-color) 65%, transparent) !important;
        }

        [data-testid="stFileUploader"] button {
            background: var(--background-color);
            border: 1px solid color-mix(in srgb, var(--text-color) 22%, transparent);
            color: var(--text-color);
        }

        [data-testid="stFileUploader"] button:hover {
            background: color-mix(in srgb, var(--secondary-background-color) 80%, var(--primary-color));
            border-color: var(--primary-color);
            color: var(--text-color);
        }

        [data-testid="stExpander"] {
            background: var(--secondary-background-color);
            border: 1px solid color-mix(in srgb, var(--text-color) 16%, transparent);
            border-radius: 10px;
            margin: 0.8rem 0 1.3rem;
        }

        [data-testid="stExpander"] summary p,
        [data-testid="stExpander"] summary span {
            color: var(--text-color) !important;
        }

        [data-testid="stFormSubmitButton"],
        [data-testid="stDownloadButton"] {
            width: 100%;
        }

        [data-testid="stFormSubmitButton"] button,
        [data-testid="stDownloadButton"] button {
            background: var(--primary-color);
            border: 1px solid var(--primary-color);
            border-radius: 9px;
            color: #ffffff;
            font-weight: 600;
            min-height: 2.8rem;
            width: 100%;
        }

        [data-testid="stFormSubmitButton"] button:hover,
        [data-testid="stDownloadButton"] button:hover {
            filter: brightness(0.92);
            color: #ffffff;
        }

        .result-note {
            background: var(--secondary-background-color);
            border: 1px solid color-mix(in srgb, var(--text-color) 16%, transparent);
            border-radius: 10px;
            color: color-mix(in srgb, var(--text-color) 72%, transparent);
            font-size: 0.9rem;
            line-height: 1.5;
            margin-top: 1.25rem;
            padding: 0.8rem 1rem;
        }

        [data-testid="stCaptionContainer"] {
            color: color-mix(in srgb, var(--text-color) 60%, transparent);
            margin-top: 1.2rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="extractor-header">
        <div class="extractor-label">Financials processing</div>
        <h1 class="extractor-title">Gain Financials XML Extractor</h1>
        <p class="extractor-description">
            Upload up to 5 XML files for annual reports from KRS. Order and file names do not matter.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
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
        label="Download filled Excel workbook",
        data=result,
        file_name="Filled_Financials.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.markdown(
        '<div class="result-note">The generated workbook is available only for '
        "this session. Download it before closing this page.</div>",
        unsafe_allow_html=True,
    )

st.caption(
    "The workbook may recalculate formulas when it is first opened in Microsoft Excel."
)
