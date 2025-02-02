import streamlit as st
import pandas as pd
import io
import zipfile
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
import time
import random

def read_excel_simple(file):
    """
    Lees een Excel-bestand dat mogelijk een foutieve styles.xml bevat.
    We openen het XLSX-bestand als een zip-archief, vervangen in xl/styles.xml 'biltinId'
    door 'builtinId', en laden vervolgens het workbook in read-only modus (zonder stijlen).
    """
    try:
        file_bytes = file.read()
        in_memory_file = BytesIO(file_bytes)
        
        # Open het XLSX-bestand als zip-archief en pas styles.xml aan indien nodig
        with zipfile.ZipFile(in_memory_file, 'r') as zin:
            out_buffer = BytesIO()
            with zipfile.ZipFile(out_buffer, 'w') as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    if item.filename == 'xl/styles.xml':
                        content = content.replace(b'biltinId', b'builtinId')
                    zout.writestr(item, content)
            out_buffer.seek(0)
        
        wb = load_workbook(out_buffer, read_only=True, data_only=True)
        ws = wb.active

        data = list(ws.values)
        if not data:
            st.error("Het Excel-bestand bevat geen data.")
            return pd.DataFrame()

        header = data[0]
        values = data[1:]
        df = pd.DataFrame(values, columns=header)
        return df
    except Exception as e:
        st.error(f"Fout bij het inlezen van de Excel: {e}")
        return pd.DataFrame()

def filter_stock(stock_file, catalog_file, progress_callback=None):
    # Lijst met grappige boodschappen (er zijn er meer dan 10)
    messages = [
        "Bezig met de producten uit te pakken...",
        "Transporteren van de producten...",
        "De producten worden gesorteerd...",
        "Producten tellen als een malle...",
        "Even een dutje voor de robots...",
        "Data wordt met precisie verwerkt...",
        "De voorraad wordt keurig gecontroleerd...",
        "Producten aan het hergroeperen...",
        "Creatieve oplossingen worden bedacht...",
        "Bijna klaar, hou vol!",
        "Laatste check: alles in de juiste doos!"
    ]
    
    # --- Stap 1: Inlezen stocklijst ---
    if progress_callback:
        progress_callback(10, random.choice(messages))
    stocklijst_df = read_excel_simple(stock_file)
    time.sleep(3)
    if stocklijst_df.empty:
        st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
        return pd.DataFrame()
        
    # --- Stap 2: Inlezen catalogus ---
    if progress_callback:
        progress_callback(30, random.choice(messages))
    try:
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()
    time.sleep(3)
    
    # --- Stap 3: Converteren en filteren van kolommen ---
    stocklijst_col = "Code"      # Of "EAN", afhankelijk van jouw data
    catalogus_col = "product_sku"
    try:
        if progress_callback:
            progress_callback(50, random.choice(messages))
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()
    time.sleep(3)
    
    try:
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]
    except Exception as e:
        st.error(f"Fout bij het filteren van data: {e}")
        return pd.DataFrame()
        
    # --- Stap 4: Samenvoegen en duplicaat verwijderen ---
    if progress_callback:
        progress_callback(70, random.choice(messages))
    try:
        merged_df = filtered_stocklijst_df.merge(
            catalogus_df[[catalogus_col]],
            left_on=stocklijst_col,
            right_on=catalogus_col,
            how="left"
        )
        # Verwijder de extra SKU-kolom (die uit de catalogus komt)
        merged_df = merged_df.drop(columns=[catalogus_col])
    except Exception as e:
        st.error(f"Fout bij het samenvoegen van data: {e}")
        return pd.DataFrame()
    time.sleep(3)
    
    # --- Stap 5: Hernoemen en afronden ---
    if progress_callback:
        progress_callback(90, random.choice(messages))
    rename_map = {
        "Code": "product_sku",
        "# stock": "product_quantity"
    }
    try:
        missing_columns = [col for col in rename_map.keys() if col not in merged_df.columns]
        if missing_columns:
            st.error(f"De volgende vereiste kolommen ontbreken in de stocklijst: {missing_columns}")
            return pd.DataFrame()
        merged_df = merged_df.rename(columns=rename_map)
        merged_df = merged_df[["product_sku", "product_quantity"]]
        merged_df["product_quantity"] = pd.to_numeric(
            merged_df["product_quantity"], errors="coerce"
        ).fillna(0).astype(int)
    except Exception as e:
        st.error(f"Fout bij het verwerken van de geëxporteerde data: {e}")
        return pd.DataFrame()
    time.sleep(3)
    
    if progress_callback:
        progress_callback(100, "De voorraad is compleet gesorteerd!")
    return merged_df

# --- Streamlit UI ---
st.title("LOE Stocklijst Filter Webapp - Door Maarten Verheyen")
st.write("Upload je stocklijst en catalogus om de gefilterde stocklijst te genereren.")

stock_file = st.file_uploader("Upload de Stocklijst uit Mercis (Excel)", type=["xls", "xlsx"])
catalog_file = st.file_uploader("Upload de Catalogus uit KMOShops (CSV)", type=["csv"])

if stock_file and catalog_file:
    if st.button("Filter Stocklijst"):
        # Maak een progress-bar en een plek voor de boodschap
        progress_bar = st.progress(0)
        message_placeholder = st.empty()
        
        # Callback om de progress-bar en boodschap bij te werken
        def progress_callback(percentage, message):
            progress_bar.progress(percentage)
            message_placeholder.text(message)
        
        with st.spinner("Even geduld, we verwerken de data..."):
            filtered_df = filter_stock(stock_file, catalog_file, progress_callback)
            
            if not filtered_df.empty:
                # Zet de DataFrame om naar CSV
                output = io.StringIO()
                filtered_df.to_csv(output, index=False, sep=';')
                output.seek(0)
                
                current_date = datetime.now().strftime("%Y-%m-%d")
                st.download_button(
                    label="Download Gefilterde Stocklijst",
                    data=output.getvalue(),
                    file_name=f"Gefilterde_Stocklijst_{current_date}.csv",
                    mime="text/csv"
                )
                st.success("De gefilterde stocklijst is succesvol gegenereerd!")
