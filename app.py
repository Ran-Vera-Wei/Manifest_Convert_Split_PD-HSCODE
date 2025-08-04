import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Manifest Converter", page_icon="📦")

st.title("📦 Manifest Converter")
st.write(
    "Upload your original manifest Excel. I’ll split **PRODUCT DESCRIPTION** and **HSCODE** into "
    "one-to-one rows (truncate to the shortest pair per row), keep all other columns, set **TOTAL QTY = 1**, "
    "and evenly distribute **WEIGHT** and **TOTAL DECLARE VALUE** per **Tracking Number**."
)

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

@st.cache_data(show_spinner=False)
def convert_manifest(file_bytes: bytes) -> pd.DataFrame:
    # Read original workbook (first sheet)
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xls, sheet_name=0)

    # Validate required columns
    required = ["Tracking Number", "PRODUCT DESCRIPTION", "HSCODE"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    # Preserve original totals by Tracking Number
    base = df[["Tracking Number"]].copy()
    # Bring WEIGHT and TOTAL DECLARE VALUE if present; coerce to numeric
    for c in ["WEIGHT", "TOTAL DECLARE VALUE"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
            base[c] = df[c]
    # Use the first non-null per Tracking Number as the original totals
    base_by_tracking = (
        base.groupby("Tracking Number")
        .agg({"WEIGHT": "first", "TOTAL DECLARE VALUE": "first"})
        .reset_index()
    )

    # Helper to split comma-separated cells
    def split_list_cell(x):
        if pd.isna(x):
            return []
        parts = [p.strip() for p in str(x).split(",")]
        return [p for p in parts if p != ""]

    # Build expanded rows (truncate to shortest pair count)
    carry_cols = df.columns.tolist()
    rows = []
    for _, r in df.iterrows():
        p_list = split_list_cell(r["PRODUCT DESCRIPTION"])
        h_list = split_list_cell(r["HSCODE"])
        n = min(len(p_list), len(h_list))
        if n == 0:
            continue
        for i in range(n):
            new_r = {c: r[c] for c in carry_cols}
            new_r["PRODUCT DESCRIPTION"] = p_list[i]
            new_r["HSCODE"] = h_list[i]
            rows.append(new_r)

    if not rows:
        return pd.DataFrame(columns=df.columns)

    expanded = pd.DataFrame(rows)

    # Ensure TOTAL QTY exists and set to 1
    if "TOTAL QTY" not in expanded.columns:
        expanded["TOTAL QTY"] = 1
    else:
        expanded["TOTAL QTY"] = 1

    # Count pairs per Tracking Number after expansion
    pair_counts = expanded.groupby("Tracking Number").size().rename("pair_count").reset_index()
    expanded = expanded.merge(pair_counts, on="Tracking Number", how="left")

    # Join the original totals
    expanded = expanded.merge(base_by_tracking, on="Tracking Number", how="left", suffixes=("", "_ORIG"))

    # Distribute WEIGHT and TOTAL DECLARE VALUE evenly across pairs
    if "WEIGHT_ORIG" in expanded.columns:
        expanded["WEIGHT"] = (pd.to_numeric(expanded["WEIGHT_ORIG"], errors="coerce") / expanded["pair_count"]).round(5)
    if "TOTAL DECLARE VALUE_ORIG" in expanded.columns:
        expanded["TOTAL DECLARE VALUE"] = (
            pd.to_numeric(expanded["TOTAL DECLARE VALUE_ORIG"], errors="coerce") / expanded["pair_count"]
        ).round(2)

    # Drop helper columns
    expanded.drop(columns=[c for c in ["WEIGHT_ORIG", "TOTAL DECLARE VALUE_ORIG", "pair_count"] if c in expanded.columns],
                  inplace=True)

    # Put key columns together near the end (optional)
    def reorder(df_in: pd.DataFrame) -> pd.DataFrame:
        cols = df_in.columns.tolist()
        keys = [c for c in ["PRODUCT DESCRIPTION", "HSCODE", "TOTAL QTY", "WEIGHT", "TOTAL DECLARE VALUE"] if c in cols]
        for c in keys:
            cols.remove(c)
        return df_in[cols + keys]
    return reorder(expanded)

if uploaded is not None:
    try:
        with st.spinner("Converting..."):
            converted = convert_manifest(uploaded.getvalue())
        st.success("Conversion complete!")
        st.dataframe(converted.head(50), use_container_width=True)

        # Prepare Excel in memory
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            converted.to_excel(writer, index=False, sheet_name="Converted")
        buf.seek(0)

        st.download_button(
            "⬇️ Download Converted Excel",
            data=buf,
            file_name="manifest_converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

st.caption("Required columns: ‘Tracking Number’, ‘PRODUCT DESCRIPTION’, and ‘HSCODE’.")
