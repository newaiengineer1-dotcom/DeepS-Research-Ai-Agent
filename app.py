import streamlit as st

from research_agent import generate_long_research_report

from report_builder import (
    create_docx,
    create_pdf,
    create_xlsx,
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Deep Research AI Agent",
    page_icon="🔎",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# PROFESSIONAL UI
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        font-size: 42px;
        font-weight: 800;
        margin-bottom: 4px;
    }

    .subtitle {
        font-size: 17px;
        opacity: 0.75;
        margin-bottom: 25px;
    }

    .metric-card {
        padding: 18px;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,0.25);
        text-align: center;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="main-title">🔎 Deep Research AI Agent</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle">'
    'Professional AI-powered research, analysis and report generation'
    '</div>',
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("⚙️ Report Settings")

    max_pages = st.slider(
        "📄 Maximum Report Pages",
        min_value=5,
        max_value=500,
        value=10,
        step=5,
    )

    st.info(
        f"""
**Selected:** {max_pages} pages

Approximate target:
**{max_pages * 500:,} words**
"""
    )

    st.markdown("---")

    st.subheader("📦 Output Formats")

    docx_enabled = st.checkbox(
        "📄 Microsoft Word / DOCX",
        value=True,
    )

    pdf_enabled = st.checkbox(
        "📕 Professional PDF",
        value=True,
    )

    xlsx_enabled = st.checkbox(
        "📊 Excel / Google Sheets",
        value=True,
    )

    markdown_enabled = st.checkbox(
        "📝 Markdown",
        value=True,
    )

    st.markdown("---")

    st.caption("🤖 AI Model")
    st.write("DeepSeek V4.1 Flash")

    st.caption("📑 Report Style")
    st.write("Professional / Consulting / Technical")


# ============================================================
# INPUT
# ============================================================

topic = st.text_area(
    "🔬 Research Topic",
    placeholder=(
        "Example:\n"
        "Global Solar PV and BESS Outlook 2026"
    ),
    height=120,
)


evidence = st.text_area(
    "📚 Additional Research Evidence / Data (Optional)",
    placeholder=(
        "Paste research notes, source information, "
        "statistics, technical information, URLs, "
        "documents or assumptions here..."
    ),
    height=180,
)


# ============================================================
# GENERATE BUTTON
# ============================================================

if st.button(
    "🚀 Generate Professional Research Report",
    type="primary",
    use_container_width=True,
):

    if not topic.strip():

        st.warning(
            "⚠️ Please enter a research topic."
        )

        st.stop()

    if not any([
        docx_enabled,
        pdf_enabled,
        xlsx_enabled,
        markdown_enabled,
    ]):

        st.warning(
            "Please select at least one output format."
        )

        st.stop()

    # --------------------------------------------------------
    # AI GENERATION
    # --------------------------------------------------------

    with st.spinner(
        f"🔎 Researching and generating approximately "
        f"{max_pages} pages..."
    ):

        try:

            report = generate_long_research_report(
                topic=topic.strip(),
                evidence=evidence.strip(),
                max_pages=max_pages,
            )

        except Exception as e:

            st.error(
                f"❌ Research generation failed: {e}"
            )

            st.stop()

    # --------------------------------------------------------
    # DOCUMENT GENERATION
    # --------------------------------------------------------

    with st.spinner(
        "📑 Building professional documents..."
    ):

        try:

            docx_data = None
            pdf_data = None
            xlsx_data = None

            if docx_enabled:

                docx_data = create_docx(
                    report,
                    topic.strip(),
                )

            if pdf_enabled:

                pdf_data = create_pdf(
                    report,
                    topic.strip(),
                )

            if xlsx_enabled:

                xlsx_data = create_xlsx(
                    report,
                    topic.strip(),
                )

        except Exception as e:

            st.error(
                f"❌ Document generation failed: {e}"
            )

            st.stop()

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    word_count = len(report.split())

    st.success(
        "✅ Professional research report generated successfully!"
    )

    # --------------------------------------------------------
    # METRICS
    # --------------------------------------------------------

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.metric(
            "Target Pages",
            max_pages,
        )

    with c2:

        st.metric(
            "Words",
            f"{word_count:,}",
        )

    with c3:

        st.metric(
            "Tables",
            len(
                report.split("|") // 100
            ) if False else "Included",
        )

    with c4:

        st.metric(
            "Format",
            "Professional",
        )

    # --------------------------------------------------------
    # DOWNLOADS
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "📥 Download Professional Report"
    )

    columns = st.columns(4)

    index = 0

    if docx_enabled:

        with columns[index]:

            st.download_button(
                "📄 Download DOCX",
                data=docx_data,
                file_name=(
                    "Professional_Research_Report.docx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                use_container_width=True,
            )

        index += 1

    if pdf_enabled:

        with columns[index]:

            st.download_button(
                "📕 Download PDF",
                data=pdf_data,
                file_name=(
                    "Professional_Research_Report.pdf"
                ),
                mime="application/pdf",
                use_container_width=True,
            )

        index += 1

    if xlsx_enabled:

        with columns[index]:

            st.download_button(
                "📊 Download XLSX",
                data=xlsx_data,
                file_name=(
                    "Research_Data.xlsx"
                ),
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                use_container_width=True,
            )

        index += 1

    if markdown_enabled:

        with columns[index]:

            st.download_button(
                "📝 Download Markdown",
                data=report,
                file_name=(
                    "Research_Report.md"
                ),
                mime="text/markdown",
                use_container_width=True,
            )

    # --------------------------------------------------------
    # PREVIEW
    # --------------------------------------------------------

    st.markdown("---")

    st.subheader(
        "👁️ Report Preview"
    )

    with st.expander(
        "Open generated report",
        expanded=True,
    ):

        st.markdown(report)
