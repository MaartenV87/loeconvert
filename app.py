import streamlit as st
import pandas as pd
import io
import zipfile
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook

def read_excel_simple(file):
    """
    Lees een Excel-bestand dat mogelijk een foutieve styles.xml bevat.
    We openen het XLSX-bestand als een zip-archief, vervangen in xl/styles.xml 'biltinId'
    door 'builtinId', en laden vervolgens het workbook in read-only modus (zonder stijlen).
    """
    try:
        # Lees de gehele inhoud van het geüploade bestand als bytes
        file_bytes = file.read()
        # Maak een in-memory stream van de bytes
        in_memory_file = BytesIO(file_bytes)
        
        # Open het XLSX-bestand als zip-archief en pas styles.xml aan indien nodig
        with zipfile.ZipFile(in_memory_file, 'r') as zin:
            out_buffer = BytesIO()
            with zipfile.ZipFile(out_buffer, 'w') as zout:
                # Loop door alle bestanden in het originele archive
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    # Als we het styles-bestand tegenkomen, vervang dan 'biltinId' door 'builtinId'
                    if item.filename == 'xl/styles.xml':
                        content = content.replace(b'biltinId', b'builtinId')
                    zout.writestr(item, content)
            # Zet de pointer aan het begin van de nieuwe stream
            out_buffer.seek(0)
        
        # Laad het workbook vanuit het aangepaste archive in read-only en data-only modus
        wb = load_workbook(out_buffer, read_only=True, data_only=True)
        ws = wb.active

        # Haal alle waarden op uit het actieve werkblad
        data = list(ws.values)
        if not data:
            st.error("Het Excel-bestand bevat geen data.")
            return pd.DataFrame()

        # Veronderstel dat de eerste rij de kolomnamen bevat
        header = data[0]
        values = data[1:]
        df = pd.DataFrame(values, columns=header)
        return df
    except Exception as e:
        st.error(f"Fout bij het inlezen van de Excel: {e}")
        return pd.DataFrame()

def filter_stock(stock_file, catalog_file):
    try:
        # Stocklijst inlezen
        stocklijst_df = read_excel_simple(stock_file)
        if stocklijst_df.empty:
            st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Fout bij het inlezen van de stocklijst: {e}")
        return pd.DataFrame()

    try:
        # Catalogus inlezen met automatische delimiter detectie
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()

    # Definieer de kolomnamen voor filtering
    stocklijst_col = "Code"      # Bijvoorbeeld: "Code" (of eventueel "EAN")
    catalogus_col = "product_sku"
    
    try:
        # Zorg dat de relevante kolommen als string worden behandeld
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()

    try:
        # Filter: behoud enkel de rijen in de stocklijst die ook in de catalogus voorkomen
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]

        # Samenvoegen: voeg de catalogus toe met een suffix zodat de kolomnaam niet conflicteert
        merged_df = filtered_stocklijst_df.merge(
            catalogus_df[[catalogus_col]],
            left_on=stocklijst_col,
            right_on=catalogus_col,
            how="left",
            suffixes=('', '_catalog')
        )
    except Exception as e:
        st.error(f"Fout bij het filteren en samenvoegen van data: {e}")
        return pd.DataFrame()

    # Hernoem kolommen en bewaar enkel de gewenste kolommen
    rename_map = {
        "Code": "product_sku",
        "# stock": "product_quantity"
    }
    try:
        # Controleer eerst of de vereiste kolommen aanwezig zijn
        missing_columns = [col for col in rename_map.keys() if col not in merged_df.columns]
        if missing_columns:
            st.error(f"De volgende vereiste kolommen ontbreken in de stocklijst: {missing_columns}")
            return pd.DataFrame()

        # Hernoem de kolom 'Code' naar 'product_sku'
        merged_df = merged_df.rename(columns=rename_map)

        # Verwijder de extra SKU-kolom uit de catalogus (die nu 'product_sku_catalog' heet)
        if 'product_sku_catalog' in merged_df.columns:
            merged_df = merged_df.drop(columns=['product_sku_catalog'])
        
        # Behoud enkel de gewenste kolommen
        merged_df = merged_df[["product_sku", "product_quantity"]]

        # Zorg dat 'product_quantity' als geheel getal wordt opgeslagen
        merged_df["product_quantity"] = pd.to_numeric(
            merged_df["product_quantity"], errors="coerce"
        ).fillna(0).astype(int)
    except Exception as e:
        st.error(f"Fout bij het verwerken van de geëxporteerde data: {e}")
        return pd.DataFrame()

    return merged_df

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
            if not filtered_df.empty:
                # Zet de DataFrame om naar een CSV-bestand
                output = io.StringIO()
                filtered_df.to_csv(output, index=False, sep=';')
                output.seek(0)
                
                # Huidige datum voor de bestandsnaam
                current_date = datetime.now().strftime("%Y-%m-%d")
                
                # Downloadknop tonen
                st.download_button(
                    label="Download Gefilterde Stocklijst",
                    data=output.getvalue(),
                    file_name=f"Gefilterde_Stocklijst_{current_date}.csv",
                    mime="text/csv"
                )
                st.success("De gefilterde stocklijst is succesvol gegenereerd!")
