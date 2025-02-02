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
    
    # Kolommen hernoemen en filteren voor export
    filtered_stocklijst_df = filtered_stocklijst_df.rename(columns={
        "Omschrijving": "product_name",
        "Code": "product_sku",
        "Verk. pr. \n€ excl.": "product_price",
        "product_weight": "product_weight",
        "product_description": "product_description"
    })
    
    # Alleen gewenste kolommen behouden
    filtered_stocklijst_df = filtered_stocklijst_df[[
        "product_name", "product_sku", "product_price", "product_weight", "product_description"
    ]]
    
    # "type" kolom toevoegen met vaste waarde "product"
    filtered_stocklijst_df.insert(0, "type", "product")
    
    return filtered_stocklijst_df

# Streamlit UI
st.title("LOE Stocklijst Filter Webapp - Door Maarten Verheyen")

st.write("Upload je stocklijst en catalogus om de gefilterde stocklijst te genereren.")

# Bestand uploads
stock_file = st.file_uploader("Upload de Stocklijst uit Mercis (Excel)", type=["xls", "xlsx"])
catalog_file = st.file_uploader("Upload de Catalogus uit KMOShops (CSV)", type=["csv"])

if stock_file and catalog_file:
    if st.button("Filter Stocklijst"):
        with st.spinner("Bezig met verwerken..."):
            filtered_df = filter_stock(stock_file, catalog_file)
            
            # Omzetten naar CSV-bestand
            output = io.StringIO()
            filtered_df.to_csv(output, index=False, sep=';')
            output.seek(0)
            
            # Zorg dat datetime correct gebruikt wordt
            from datetime import datetime
            current_date = datetime.now().strftime("%Y-%m-%d")
            
            # Download knop tonen
            st.download_button(
                label="Download Gefilterde Stocklijst",
                data=output.getvalue(),
                file_name=f"Gefilterde_Stocklijst_{current_date}.csv",
                mime="text/csv"
            )
            
            st.success("De gefilterde stocklijst is succesvol gegenereerd!")
