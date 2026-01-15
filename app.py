import streamlit as st
import pandas as pd
import numpy as np
import io

# -----------------------------------------------------------
# Load LTL Qty file (pre-loaded in your system)
# -----------------------------------------------------------
LTL_QTY_PATH = "LTL_qty.xlsx"

@st.cache_data
def load_ltl_qty():
    return pd.read_excel(LTL_QTY_PATH)

# -----------------------------------------------------------
# Process uploaded SAP Export files
# -----------------------------------------------------------
def process_order_export(files, ltl_qty_df):

    # Read and combine uploaded Excel files
    dfs = []
    for file in files:
        df = pd.read_excel(file)
        dfs.append(df)

    df_order_export = pd.concat(dfs, ignore_index=True)

    # Merge with LTL Qty file
    df_orders = pd.merge(
        df_order_export,
        ltl_qty_df[['SAP Code', 'LTL Qty', 'Case_Pallet']],
        left_on='Material',
        right_on='SAP Code',
        how='inner'
    )

    df_LTL = df_orders.copy()

    # Filter LTL orders -- needs to be done after grouping by PO
    #df_LTL = df_orders[
        #(df_orders['Order Quantity'] >= df_orders['LTL Qty']) |
        #(df_orders['LTL Qty'].isna() == True)
    #]

    # Convert weight to Pounds
    df_LTL['Gross weight'] = df_LTL['Gross weight'] * 2.20462

    # Base columns for cleanup
    columns = [
        'Purchase order no.',
        'Sales document',
        'Material',
        'Order Quantity',
        'Gross weight',
        'Case_Pallet',
        'LTL Qty'
    ]

    df_LTL_clean = df_LTL[columns].copy()
    df_LTL_clean['DN'] = ""

    # Grouping logic
    df_LTL_grouped = (
        df_LTL_clean
        .groupby('Purchase order no.', as_index=False)
        .agg({
            'Order Quantity': 'sum',
            'Gross weight': 'sum',
            'Case_Pallet': 'min',
            'LTL Qty': 'min',
            **{
                col: 'first'
                for col in df_LTL_clean.columns
                if col not in ['Purchase order no.', 'Order Quantity', 'Gross weight']
            }
        })
    )

     # Filter LTL orders
    df_LTL_grouped = df_LTL_grouped[
        (df_LTL_grouped['Order Quantity'] >= df_LTL_grouped['LTL Qty']) |
        (df_LTL_grouped['LTL Qty'].isna() == True)
    ]

    # Drop LTL Qty columns
    df_LTL_grouped = df_LTL_grouped.drop(columns=['LTL Qty'])

    # Pallet quantity
    df_LTL_grouped['Pallet_qty'] = np.ceil(
        df_LTL_grouped['Order Quantity'] / df_LTL_grouped['Case_Pallet']
    )

    return df_LTL_grouped


# -----------------------------------------------------------
# UI – Streamlit App
# -----------------------------------------------------------
st.set_page_config(page_title="LTL Order Cleaner", layout="centered")

st.title("📦 LTL Order Cleaning Tool")
st.write("Upload SAP Order Export file(s) and automatically generate the cleaned LTL grouping sheet.")

st.markdown("---")

# Load LTL Qty file
try:
    ltl_qty_df = load_ltl_qty()
    st.success("LTL Qty file loaded successfully.")
except Exception as e:
    st.error(f"❌ Error loading LTL_qty.xlsx: {e}")
    st.stop()

# Upload area
uploaded_files = st.file_uploader(
    "📤 Upload SAP Order Export Excel file(s)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:

    if st.button("▶️ Process Files"):
        try:
            df_output = process_order_export(uploaded_files, ltl_qty_df)

            st.success("Processing completed!")
            st.dataframe(df_output.head(50))

            # Prepare file for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                df_output.to_excel(writer, index=False, sheet_name="LTL_Output")
            excel_data = buffer.getvalue()

            st.download_button(
                "⬇️ Download Cleaned LTL File",
                data=excel_data,
                file_name="LTL_Cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ Error processing files: {e}")

else:
    st.info("Upload SAP Order Export Excel files to begin.")


