import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Manifest Converter", page_icon="📦")

st.title("📦 Manifest Converter")
st.write(
    "Upload your original manifest Excel. This tool will split `PRODUCT DESCRIPTION` and `HSCODE` into one-to-one rows, "
    "set `TOTAL QTY = 1`, evenly distribute `WEIGHT` and `TOTAL DECLARE VALUE` per `Tracking Number`, and export to your final template format."
)

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

@st.cache_data(show_spinner=False)
def convert_manifest_to_template(file_bytes: bytes) -> pd.DataFrame:
    # Step 1: Read original
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    df = pd.read_excel(xls, sheet_name=0)

    # Step 2: Explode PRODUCT DESCRIPTION and HSCODE (truncate to shortest)
    df["PRODUCT DESCRIPTION"] = df["PRODUCT DESCRIPTION"].astype(str).str.split(",")
    df["HSCODE"] = df["HSCODE"].astype(str).str.split(",")
    min_len = [min(len(p), len(h)) for p, h in zip(df["PRODUCT DESCRIPTION"], df["HSCODE"])]
    df["PRODUCT DESCRIPTION"] = [p[:n] for p, n in zip(df["PRODUCT DESCRIPTION"], min_len)]
    df["HSCODE"] = [h[:n] for h, n in zip(df["HSCODE"], min_len)]
    df_expanded = df.loc[df.index.repeat(min_len)].copy()
    df_expanded["PRODUCT DESCRIPTION"] = [i for sub in df["PRODUCT DESCRIPTION"] for i in sub]
    df_expanded["HSCODE"] = [i for sub in df["HSCODE"] for i in sub]
    df_expanded["TOTAL QTY"] = 1

    # Step 3: Distribute WEIGHT and DECLARE VALUE evenly per tracking number
    for col in ["WEIGHT", "TOTAL DECLARE VALUE"]:
        if col in df_expanded.columns:
            df_expanded[col] = pd.to_numeric(df_expanded[col], errors="coerce")
    counts = df_expanded.groupby("Tracking Number").size().rename("_count")
    df_expanded = df_expanded.merge(counts, on="Tracking Number")
    df_expanded["WEIGHT"] = df_expanded.groupby("Tracking Number")["WEIGHT"].transform("first") / df_expanded["_count"]
    df_expanded["TOTAL DECLARE VALUE"] = df_expanded.groupby("Tracking Number")["TOTAL DECLARE VALUE"].transform("first") / df_expanded["_count"]
    df_expanded.drop(columns="_count", inplace=True)

    # Step 4: Map to template columns
    province_abbr = {
        "Beijing": "BJ", "Tianjin": "TJ", "Shanghai": "SH", "Chongqing": "CQ",
        "Guangdong": "GD", "Guangxi": "GX", "Guizhou": "GZ", "Yunnan": "YN", "Hainan": "HI",
        "Sichuan": "SC", "Jiangsu": "JS", "Zhejiang": "ZJ", "Anhui": "AH", "Fujian": "FJ",
        "Jiangxi": "JX", "Shandong": "SD", "Henan": "HA", "Hubei": "HB", "Hunan": "HN",
        "Hebei": "HE", "Shanxi": "SX", "Inner Mongolia": "NM", "Shaanxi": "SN", "Gansu": "GS",
        "Qinghai": "QH", "Ningxia": "NX", "Xinjiang": "XJ", "Tibet": "XZ", "Heilongjiang": "HL",
        "Jilin": "JL", "Liaoning": "LN"
    }
    city_to_province = {
        "Guangzhou": "Guangdong", "Shenzhen": "Guangdong", "Foshan": "Guangdong",
        "Beijing": "Beijing", "Shanghai": "Shanghai", "Tianjin": "Tianjin",
        "Chengdu": "Sichuan", "Wuhan": "Hubei", "Hangzhou": "Zhejiang", "Nanjing": "Jiangsu"
    }
    template_cols = [
        "consignor_item_id", "display_id", "receptacle_id", "tracking_number", "sender_name",
        "sender_orgname", "sender_address1", "sender_address2", "sender_district", "sender_city",
        "sender_state", "sender_zip5", "sender_zip4", "sender_country", "sender_phone", "sender_email",
        "sender_url", "recipient_name", "recipient_orgname", "recipient_address1", "recipient_address2",
        "recipient_district", "recipient_city", "recipient_state", "recipient_zip5", "recipient_zip4",
        "recipient_country", "recipient_phone", "recipient_email", "recipient_addr_type", "return_name",
        "return_orgname", "return_address1", "return_address2", "return_district", "return_city",
        "return_state", "return_zip5", "return_zip4", "return_country", "return_phone", "return_email",
        "mail_type", "pieces", "weight", "length", "width", "height", "girth", "value", "machinable",
        "po_box_flag", "gift_flag", "commercial_flag", "customs_quantity_units", "dutiable",
        "duty_pay_by", "product", "description", "url", "sku", "country_of_origin", "manufacturer",
        "harmonization_code", "unit_value", "quantity", "total_value", "total_weight"
    ]
    col_map = {
        "receptacle_id": "Bag ID", "consignor_item_id": "BG Number", "tracking_number": "Tracking Number",
        "sender_name": "SHIPPER", "sender_address1": "SHIPPER ADDRESS", "sender_city": "CITY NAME SHIPPER",
        "sender_country": "COUNTRY CODE SHIPPER", "recipient_name": "Consignee Name",
        "recipient_address1": "Consignee Address", "recipient_city": "Consignee City",
        "recipient_state": "Consignee Province", "recipient_zip5": "Consignee Post Code",
        "recipient_country": "Country of Destination", "weight": "WEIGHT", "value": "TOTAL DECLARE VALUE",
        "description": "PRODUCT DESCRIPTION", "harmonization_code": "HSCODE", "unit_value": "TOTAL DECLARE VALUE",
        "total_value": "TOTAL DECLARE VALUE", "total_weight": "WEIGHT"
    }
    result = pd.DataFrame(columns=template_cols)
    for col in template_cols:
        src = col_map.get(col)
        if src:
            result[col] = df_expanded[src]
        elif col in ["pieces", "length", "width", "height", "girth", "quantity"]:
            result[col] = 1
        else:
            result[col] = ""
    result["sender_state"] = result["sender_city"].map(city_to_province).map(province_abbr)
    return result

if uploaded:
    try:
        df_final = convert_manifest_to_template(uploaded.getvalue())
        st.success("Conversion complete!")
        st.dataframe(df_final.head(50), use_container_width=True)

        # Prepare download
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False)
        buffer.seek(0)

        st.download_button(
            label="📂 Download Converted Template",
            data=buffer,
            file_name="manifest_converted_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Error: {e}")
