import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Manifest Converter", page_icon="📦")

st.title("📦 Manifest Converter")
st.write(
    "Upload your original manifest Excel. "
    "I'll split `PRODUCT DESCRIPTION` and `HSCODE` into one-to-one rows, "
    "keep other columns, set `TOTAL QTY = 1`, and distribute `WEIGHT` and "
    "`TOTAL DECLARE VALUE` evenly by `Tracking Number`."
)

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

@st.cache_data(show_spinner=False)
def convert_manifest(file_bytes: bytes) -> pd.DataFrame:
    # Read original
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xls, sheet_name=0)

    # Ensure expected columns exist
    required_cols = ["Tracking Number", "PRODUCT DESCRIPTION", "HSCODE"]
    for c in required_cols:
        if c not in df.columns:
            raise ValueError(f"Missing required column: '{c}'")

    # Keep originals for per-tracking aggregates
    # (If duplicates per tracking exist, take the first non-null numeric)
    originals = (
        df[["Tracking Number", "WEIGHT", "TOTAL DECLARE VALUE"]]
        .copy()
    )

    # Coerce numeric columns
    for col in ["WEIGHT", "TOTAL DECLARE VALUE"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    def split_list_cell(x):
        if pd.isna(x):
            return []
        return [part.strip() for part in str(x).split(",") if str(part).strip() != ""]

    # Build expanded rows (truncate to shortest pair length)
    expanded_rows = []
    # capture columns to carry over
    carry_cols = list(df.columns)

    for _, row in df.iterrows():
        p_list = split_list_cell(row["PRODUCT DESCRIPTION"])
        h_list = split_list_cell(row["HSCODE"])
        n = min(len(p_list), len(h_list))

        if n == 0:
            # No pairs; keep nothing (or keep 1 row? Here we skip)
            continue

        for i in range(n):
            new_row = {col: row[col] for col in carry_cols}
            new_row["PRODUCT DESCRIPTION"] = p_list[i]
            new_row["HSCODE"] = h_list[i]
            expanded_rows.append(new_row)

    if not expanded_rows:
        return pd.DataFrame(columns=df.columns)

    expanded = pd.DataFrame(expanded_rows)

    # TOTAL QTY = 1 for every row
    if "TOTAL QTY" in expanded.columns:
        expanded["TOTAL QTY"] = 1
    else:
        expanded.insert(len(expanded.columns), "TOTAL QTY", 1)

    # Compute per-tracking counts after expansion
    counts_per_tracking = expanded.groupby("Tracking Number").size().rename("pair_count")
    expanded = expanded.merge(counts_per_tracking, on="Tracking Number", how="left")

    # Bring original WEIGHT & VALUE per tracking (first non-null values)
    base_by_tracking = (
        originals.groupby("Tracking Number")
        .agg({"WEIGHT": "first", "TOTAL DECLARE VALUE": "first"})
        .reset_index()
    )
    expanded = expanded.merge(base_by_tracking, on="Tracking Number", how="left", suffixes=("", "_ORIG"))

    # Distribute WEIGHT and TOTAL DECLARE VALUE evenly by Tracking Number
    if "WEIGHT_ORIG" in expanded.columns and "pair_count" in expanded.columns:
        expanded["WEIGHT"] = (
            pd.to_numeric(expanded["WEIGHT_ORIG"], errors="coerce") / expanded["pair_count"]
        ).round(5)

    if "TOTAL DECLARE VALUE_ORIG" in expanded.columns and "pair_count" in expanded.columns:
        expanded["TOTAL DECLARE VALUE"] = (
            pd.to_numeric(expanded["TOTAL DECLARE VALUE_ORIG"], errors="coerce") / expanded["pair_count"]
        ).round(2)

    # Clean helper columns
    expanded.drop(columns=[c for c in ["WEIGHT_ORIG", "TOTAL DECLARE VALUE_ORIG", "pair_count"] if c in expanded.columns],
                  inplace=True)

    # Reorder columns: move the three key columns next to each other if present
    def reorder_cols(df_in: pd.DataFrame) -> pd.DataFrame:
        cols = df_in.columns.tolist()
        ordered = []
        for c in ["PRODUCT DESCRIPTION", "HSCODE", "TOTAL QTY", "WEIGHT", "TOTAL DECLARE VALUE"]:
            if c in cols:
                ordered.append(c)
                cols.remove(c)
        return df_in[cols[:]] if not ordered else df_in[[*cols, *ordered]]
    expanded = reorder_cols(expanded)

    return expanded

if uploaded is not None:
    try:
        with st.spinner("Converting..."):
            converted_df = convert_manifest(uploaded.getvalue())

        st.success("Conversion complete!")

        st.dataframe(converted_df.head(50), use_container_width=True)

        # Prepare Excel in-memory
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            converted_df.to_excel(writer, index=False, sheet_name="Converted")
        output.seek(0)

        st.download_button(
            label="⬇️ Download Converted Excel",
            data=output,
            file_name="manifest_converted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Error: {e}")
        st.stop()

st.caption("Tip: columns must include ‘Tracking Number’, ‘PRODUCT DESCRIPTION’, and ‘HSCODE’.")

