import io
from pathlib import Path

import streamlit as st

from core.settings import DEFAULT_EXCLUDE, SIDEBAR_CSS
from core.types import ExtractOptions
from core.extract import process_url
from core.docx_builder import build_docx
from core.utils import safe_filename

# -------------------------
# STREAMLIT APP
# -------------------------

icon_path = Path("assets/JAFavicon.png")
st.set_page_config(
    page_title="Content Rec Template Tool",
    page_icon=str(icon_path) if icon_path.exists() else None,
    layout="wide",
)

st.markdown(SIDEBAR_CSS, unsafe_allow_html=True)
st.title("Content Rec Template Generation Tool")

# session state for stable downloads across reruns
if "single_docx" not in st.session_state:
    st.session_state.single_docx = None
    st.session_state.single_docx_name = None
if "batch_zip" not in st.session_state:
    st.session_state.batch_zip = None
    st.session_state.batch_zip_name = "content_recommendations.zip"

with st.sidebar:
    st.header("Template & Options")
    tpl_file = st.file_uploader("Upload Template as .DOCX file", type=["docx"])
    st.caption("This should be your blank template with placeholders (e.g., [PAGE], [DATE], [PAGE BODY CONTENT], etc.).")

    st.divider()
    st.subheader("Need a template?")
    template_path = Path("assets/blank_template.docx")
    if template_path.exists():
        with open(template_path, "rb") as file:
            st.download_button(
                label="Download a Blank Template",
                data=file,
                file_name="blank_template.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.info("Place your template at assets/blank_template.docx to enable this download.")

    st.caption("Once downloaded, you'll still need to upload this above, but this version is a decent starting point.")

    st.divider()
    st.subheader("Exclude Selectors")
    exclude_txt = st.text_area(
        "Comma-separated CSS selectors to remove from <body>",
        value=", ".join(DEFAULT_EXCLUDE),
        height=120,
    )
    exclude_selectors = [s.strip() for s in exclude_txt.split(",") if s.strip()]

    st.subheader("Link formatting")
    annotate_links = st.toggle("Append (→ URL) after anchor text", value=False)

    # Toggle instead of checkbox for consistency with Link formatting
    remove_before_h1 = st.toggle("Delete everything before first <h1>", value=False)

    # Include <img> src alongside alt text
    include_img_src = st.toggle("Include <img> src in output", value=False)

    st.caption("Timezone fixed to Europe/London; dates in DD/MM/YYYY.")

# --- Tabs ---
tab1, tab2 = st.tabs(["Single URL", "Batch (CSV)"])

with tab1:
    st.subheader("Single page")

    # Agency / Client fields just above the URL field
    col0a, col0b = st.columns([1, 1])
    with col0a:
        agency_name = st.text_input("Agency Name", value="", placeholder="e.g., JA Consulting")
    with col0b:
        client_name = st.text_input("Client Name", value="", placeholder="e.g., Workspace")

    url = st.text_input("URL", value="https://www.example.com")

    col_a, col_b = st.columns([1, 1])
    with col_a:
        do_preview = st.button("Extract preview")
    with col_b:
        do_doc = st.button("Generate DOCX")

    if do_preview or do_doc:
        if not tpl_file and do_doc:
            st.error("Please upload your Rec Template.docx in the sidebar first.")
        else:
            try:
                opts = ExtractOptions(
                    exclude_selectors=exclude_selectors,
                    annotate_links=annotate_links,
                    remove_before_h1=remove_before_h1,
                    include_img_src=include_img_src,
                )
                meta, lines = process_url(url, opts)

                # Inject Agency/Client into meta for downstream use
                meta["agency"] = agency_name.strip()
                meta["client_name"] = client_name.strip()

                st.success("Extracted successfully.")
                with st.expander("Meta (preview)", expanded=True):
                    st.write(meta)
                with st.expander("Signposted content (preview)", expanded=True):
                    st.text("\n".join(lines))

                if do_doc:
                    out_bytes = build_docx(tpl_file.read(), meta, lines)
                    fname = safe_filename(f"{meta['page']} - Content Recommendations") + ".docx"
                    # store for stable download across reruns
                    st.session_state.single_docx = out_bytes
                    st.session_state.single_docx_name = fname
            except Exception as e:
                st.exception(e)

    # render download button if we have a generated file
    if st.session_state.single_docx:
        st.download_button(
            "Download DOCX",
            data=st.session_state.single_docx,
            file_name=st.session_state.single_docx_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="dl_single_docx",
        )

with tab2:
    import csv
    import zipfile

    st.subheader("Batch process CSV")
    st.caption("Upload a CSV with a header row; required column: url. Optional: out_name.")
    batch_file = st.file_uploader("CSV file", type=["csv"], key="csv")
    if st.button("Run batch"):
        if not tpl_file:
            st.error("Please upload your Rec Template.docx in the sidebar first.")
        elif not batch_file:
            st.error("Please upload a CSV.")
        else:
            tpl_bytes = tpl_file.read()
            rows = list(csv.DictReader(io.StringIO(batch_file.getvalue().decode("utf-8"))))
            if not rows:
                st.error("CSV appears empty.")
            elif "url" not in rows[0]:
                st.error("CSV must include a 'url' column.")
            else:
                memzip = io.BytesIO()
                zf = zipfile.ZipFile(memzip, "w", zipfile.ZIP_DEFLATED)
                results = []
                opts = ExtractOptions(
                    exclude_selectors=exclude_selectors,
                    annotate_links=annotate_links,
                    remove_before_h1=remove_before_h1,
                    include_img_src=include_img_src,
                )
                for i, row in enumerate(rows, 1):
                    u = row["url"].strip()
                    try:
                        meta, lines = process_url(u, opts)
                        out_name_raw = (row.get("out_name") or f"{meta['page']} - Content Recommendations").strip()
                        out_name = safe_filename(out_name_raw)
                        out_bytes = build_docx(tpl_bytes, meta, lines)
                        zf.writestr(f"{out_name}.docx", out_bytes)
                        results.append({"url": u, "status": "ok", "file": f"{out_name}.docx"})
                    except Exception as e:
                        results.append({"url": u, "status": f"error: {e}", "file": ""})
                zf.close()
                memzip.seek(0)
                st.success("Batch complete.")
                st.dataframe(results)

                # store for stable download across reruns
                st.session_state.batch_zip = memzip.read()

    # render batch download if available
    if st.session_state.batch_zip:
        st.download_button(
            "Download ZIP",
            data=st.session_state.batch_zip,
            file_name=st.session_state.batch_zip_name,
            mime="application/zip",
            key="dl_batch_zip",
        )
