# Manifest_Convert_Split_PD-HSCODE
Convertthe origin file to the one-to-one format (splitting PRODUCT DESCRIPTION and HSCODE, truncating to the shortest pair per row, keeping other columns the same, setting TOTAL QTY = 1, and distributing WEIGHT and TOTAL DECLARE VALUE evenly per Tracking Number)

# Requirements
streamlit
pandas
openpyxl
