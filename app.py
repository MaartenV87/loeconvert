import streamlit as st
import pandas as pd
import io
import zipfile
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
import base64

# --- Pagina-configuratie en aangepaste CSS ---
st.set_page_config(page_title="LOE Stocklijst Filter", page_icon=":package:", layout="wide")

st.markdown("""
    <style>
    /* Algemene pagina styling */
    .reportview-container {
        background: #f0f2f6;
    }
    /* Sidebar styling */
    .sidebar .sidebar-content {
        background: #ffffff;
    }
    /* Banner en header styling */
    .header-banner {
        text-align: center;
        padding: 20px 0;
    }
    .header-banner h1 {
        font-size: 3em;
        color: #333333;
        margin: 0;
    }
    .header-banner h3 {
        font-size: 1em;
        color: #555555;
        margin: 0;
    }
    /* Centraal uitgelijnde knoppen */
    .centered-button {
        display: flex;
        justify-content: center;
        margin-top: 20px;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Banner ---
# (Pas de URL hieronder aan naar een eigen afbeelding indien gewenst)
st.image("https://via.placeholder.com/1200x300?text=LOE+Stocklijst+Filter", use_container_width=True)
st.markdown("<div class='header-banner'><h1>LOE Stocklijst Filter App</h1><h3>Gemaakt door Maarten Verheyen</h3></div>", unsafe_allow_html=True)

# --- Sidebar met instructies ---
st.sidebar.title("Instructies")
st.sidebar.info(
    """
    **Stap 1:** Upload de *Stocklijst uit Mercis (Excel)* in de linkerkolom.  
    **Stap 2:** Upload de *Catalogus uit KMOShops (CSV)* in de rechterkolom.  
    **Stap 3:** Klik op **Filter Stocklijst** om de verwerking te starten.  
    **Stap 4:** Download de gefilterde stocklijst zodra de verwerking voltooid is.
    """
)

# --- Functies voor inlezen en verwerken van de data ---

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
    # Stap 1: Inlezen van de stocklijst
    if progress_callback:
        progress_callback(10)
    stocklijst_df = read_excel_simple(stock_file)
    if stocklijst_df.empty:
        st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
        return pd.DataFrame()

    # Stap 2: Inlezen van de catalogus
    if progress_callback:
        progress_callback(30)
    try:
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()

    # Kolomnamen voor filtering
    stocklijst_col = "Code"      # Bijvoorbeeld: "Code" (of "EAN")
    catalogus_col = "product_sku"
    
    # Stap 3: Converteren van kolommen naar string
    if progress_callback:
        progress_callback(50)
    try:
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()

    # Stap 4: Filteren van de stocklijst op basis van de catalogus
    try:
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]
    except Exception as e:
        st.error(f"Fout bij het filteren van data: {e}")
        return pd.DataFrame()

    # Stap 5: Samenvoegen en duplicaat verwijderen
    if progress_callback:
        progress_callback(70)
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

    # Stap 6: Hernoemen van kolommen en afronden
    if progress_callback:
        progress_callback(90)
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

    if progress_callback:
        progress_callback(100)
    return merged_df

# --- Bestandsuploads in twee kolommen ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("Stocklijst uit Mercis (Excel)")
    stock_file = st.file_uploader("Upload hier je Stocklijst", type=["xls", "xlsx"])

with col2:
    st.subheader("Catalogus uit KMOShops (CSV)")
    catalog_file = st.file_uploader("Upload hier je Catalogus", type=["csv"])

# --- Verwerking en download ---
if stock_file and catalog_file:
    # Centreer de Filter-knop met behulp van drie kolommen (button in de middelste kolom)
    cols = st.columns(3)
    with cols[1]:
        if st.button("Filter Stocklijst"):
            progress_bar = st.progress(0)
            def progress_callback(percentage):
                progress_bar.progress(percentage)
            with st.spinner("Bezig met verwerken..."):
                filtered_df = filter_stock(stock_file, catalog_file, progress_callback)
            if not filtered_df.empty:
                output = io.StringIO()
                filtered_df.to_csv(output, index=False, sep=';')
                output.seek(0)
                current_date = datetime.now().strftime("%Y-%m-%d")
                csv_data = output.getvalue()
                # Base64 encode van de CSV-data
                b64 = base64.b64encode(csv_data.encode()).decode()
                download_link = f"""
                <div style="text-align: center; margin-top: 20px;">
                    <a href="data:file/csv;base64,{b64}" download="Gefilterde_Stocklijst_{current_date}.csv"
                    style="background-color: #28a745; color: white; padding: 10px 20px; border-radius: 5px; text-decoration: none; font-weight: bold;">
                        Download Gefilterde Stocklijst
                    </a>
                </div>
                """
                st.markdown(download_link, unsafe_allow_html=True)
                st.success("De gefilterde stocklijst is succesvol gegenereerd!")
