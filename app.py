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
    Het bestand wordt als zip-archief geopend, waarbij in xl/styles.xml 'biltinId'
    wordt vervangen door 'builtinId'. Daarna wordt het workbook in read-only modus
    ingelezen (zonder stijlen).
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

def filter_stock(stock_file, catalog_file):
    # Stap 1: Inlezen van de stocklijst
    stocklijst_df = read_excel_simple(stock_file)
    if stocklijst_df.empty:
        st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
        return pd.DataFrame()

    # Stap 2: Inlezen van de catalogus
    try:
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()

    # Kolomnamen voor filtering
    stocklijst_col = "Code"      # Bijvoorbeeld: "Code" of "EAN"
    catalogus_col = "product_sku"
    
    # Stap 3: Converteren van de kolommen naar string
    try:
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()

    # Stap 4: Filteren op rijen die in de catalogus voorkomen
    try:
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]
    except Exception as e:
        st.error(f"Fout bij het filteren van data: {e}")
        return pd.DataFrame()

    # Stap 5: Samenvoegen en duplicaat verwijderen
    try:
        merged_df = filtered_stocklijst_df.merge(
            catalogus_df[[catalogus_col]],
            left_on=stocklijst_col,
            right_on=catalogus_col,
            how="left"
        )
        # Verwijder de extra SKU-kolom (uit de catalogus)
        merged_df = merged_df.drop(columns=[catalogus_col])
    except Exception as e:
        st.error(f"Fout bij het samenvoegen van data: {e}")
        return pd.DataFrame()

    # Stap 6: Hernoemen en afronden
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

    return merged_df

# --- Streamlit UI ---

st.title("LOE Stocklijst Filter Webapp - Door Maarten Verheyen")
st.write("Upload hieronder je bestanden om de gefilterde stocklijst te genereren.")

# Gebruik twee kolommen voor een duidelijk onderscheid tussen de uploaders
col1, col2 = st.columns(2)

with col1:
    st.header("Stocklijst uit Mercis (Excel)")
    stock_file = st.file_uploader("Upload hier je Stocklijst", type=["xls", "xlsx"])

with col2:
    st.header("Catalogus uit KMOShops (CSV)")
    catalog_file = st.file_uploader("Upload hier je Catalogus", type=["csv"])

if stock_file and catalog_file:
    if st.button("Filter Stocklijst"):
        with st.spinner("Even geduld, we verwerken de data..."):
            filtered_df = filter_stock(stock_file, catalog_file)
            if not filtered_df.empty:
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
