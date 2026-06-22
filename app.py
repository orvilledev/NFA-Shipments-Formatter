import streamlit as st

from formatter import format_shipment

st.set_page_config(page_title="NFA Shipments Formatter", page_icon="📦", layout="centered")

st.title("NFA Shipments Formatter")
st.markdown(
    "Upload an **FBA Carton Detail** workbook to generate a formatted "
    "**Box Contents** file."
)

uploaded = st.file_uploader(
    "Upload FBA Carton Detail file",
    type=["xlsx", "xlsm", "xltx", "xltm", "xls", "csv", "tsv"],
    help="Supported formats: .xlsx, .xlsm, .xltx, .xltm, .xls, .csv, .tsv. "
    "The file should contain 'FBA Carton Detail' in cell A1.",
)

if uploaded is not None:
    try:
        output_buffer, filename, summary = format_shipment(uploaded, uploaded.name)

        st.success("File processed successfully.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Shipment ID", summary["shipment_id"])
        col2.metric("Boxes", summary["carton_count"])
        col3.metric("Line Items", summary["line_count"])
        col4.metric("Total Units", summary["total_units"])

        st.download_button(
            label=f"Download {filename}",
            data=output_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander("Output details"):
            st.markdown(
                f"""
                **Output filename:** `{filename}`

                **Sheets:**
                - `Box Contents` — Box Number (number), UPC (text), Qty (number)
                - `Weights and Dimensions` — Box Number (number), Weight (number),
                  Carton Length/Width/Height (numbers)
                """
            )

    except Exception as exc:
        st.error(f"Failed to process file: {exc}")

else:
    st.info("Upload an Excel or CSV file to get started.")
