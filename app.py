import streamlit as st
import pandas as pd
import io
from datetime import datetime

def filter_stock(stock_file, catalog_file):
    # Stocklijst inlezen
    stocklijst_df = pd.read_excel(stock_file)
    
    # Catalogus inlezen met automatische delimiter detectie
    catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    
    # Kolommen identificeren voor filtering
    stocklijst_col = "Code"  # Alternatief: "EAN"
    catalogus_col = "product_sku"
    
    # Converteren naar string om mogelijke datatypeverschillen te voorkomen
    stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
    catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    
    # Filteren: Alleen rijen uit de stocklijst behouden die in de catalogus staan
    filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]
    
    return filtered_stocklijst_df

# Streamlit UI
st.title("LOE Stocklijst Filter Webapp - Door Maarten Verheyen")

st.write("Upload je stocklijst en catalogus om de gefilterde stocklijst te genereren.")

# Bestand uploads
stock_file = st.file_uploader("Upload de Stocklijst uit Mercis (Excel)", type=["xls", "xlsx"])
catalog_file = st.file_uploader("Upload de Catalogus uit KMOShops (CSV)", type=["csv"])

if stock_file and catalog_file:
    if st.button("Filter Stocklijst"):
        filtered_df = filter_stock(stock_file, catalog_file)
        
        # Omzetten naar Excel-bestand
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            filtered_df.to_excel(writer, index=False)
        output.seek(0)
        
        # Zorg dat datetime correct gebruikt wordt
        from datetime import datetime
        current_date = datetime.now().strftime("%Y-%m-%d")
        
        # Download knop tonen
        st.download_button(
            label="Download Gefilterde Stocklijst",
            data=output,
            file_name=f"Gefilterde_Stocklijst_{current_date}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("De gefilterde stocklijst is succesvol gegenereerd!")
