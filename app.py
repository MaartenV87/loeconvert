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
    /* Header styling (alleen tekst, geen bannerafbeelding) */
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
    /* Centraal uitgelijnde knop */
    .centered-button > div.stButton {
        display: flex;
        justify-content: center;
    }
    </style>
    """, unsafe_allow_html=True)

# --- Header (alleen tekst) ---
st.markdown("<div class='header-banner'><h1>LOE Stocklijst Filter App</h1><h3>Gemaakt door Maarten Verheyen</h3></div>", unsafe_allow_html=True)

# --- Sidebar met instructies ---
st.sidebar.title("Instructies")
st.sidebar.info(
    """
    **Stap 1:** Upload de *Stocklijst uit Mercis (Excel)* in de linkerkolom.  
    **Stap 2:** Upload de *Catalogus uit KMOShops (CSV)* in de rechterkolom.  
    **Stap 3:** Klik op **Filter Stocklijst** om de verwerking te starten.  
    **Stap 4:** Download de gefilterde stocklijst en bekijk het verschiloverzicht onderaan.
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

    # Stap 2: Inlezen van de catalogus (voor de filtering)
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

# --- Centraal geplaatste Filter-knop ---
st.markdown('<div class="centered-button">', unsafe_allow_html=True)
if st.button("Filter Stocklijst"):
    progress_bar = st.progress(0)
    def progress_callback(percentage):
        progress_bar.progress(percentage)
    with st.spinner("Bezig met verwerken..."):
        filtered_df = filter_stock(stock_file, catalog_file, progress_callback)
    if not filtered_df.empty:
        # --- Excel Export (ongewijzigd) ---
        output = io.StringIO()
        filtered_df.to_csv(output, index=False, sep=';')
        output.seek(0)
        current_date = datetime.now().strftime("%Y-%m-%d")
        csv_data = output.getvalue()
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
        
        # --- Overzicht van Verschillen ---
        try:
            # Reset de pointer van catalog_file voordat we deze opnieuw inlezen
            catalog_file.seek(0)
            # Probeer eerst de catalogus in te lezen met automatische delimiter detectie
            try:
                catalogus_df_full = pd.read_csv(catalog_file, sep=None, engine="python")
            except Exception as e:
                # Fallback: probeer met een komma
                try:
                    catalogus_df_full = pd.read_csv(catalog_file, delimiter=',')
                except Exception as e:
                    # Fallback: probeer met een puntkomma
                    catalogus_df_full = pd.read_csv(catalog_file, delimiter=';')
            catalogus_df_full['product_sku'] = catalogus_df_full['product_sku'].astype(str)
            
            # Maak een kopie van de export en hernoem de stocklist-hoeveelheid naar 'Nieuw aantal'
            filtered_export = filtered_df.copy().rename(columns={'product_quantity': 'Nieuw aantal'})
            # Voeg de catalogus-informatie toe: 'Omschrijving' en 'product_quantity' als 'Vorig aantal'
            diff_df = pd.merge(filtered_export, 
                               catalogus_df_full[['product_sku', 'Omschrijving', 'product_quantity']], 
                               on='product_sku', how='left')
            diff_df = diff_df.rename(columns={'product_quantity': 'Vorig aantal'})
            # Bereken het verschil: (Nieuw aantal - Vorig aantal)
            diff_df['Verschil'] = diff_df['Nieuw aantal'] - diff_df['Vorig aantal']
            # Houd enkel producten met een verschil
            diff_df = diff_df[diff_df['Verschil'] != 0]
            # Behoud de gewenste kolommen en herschik de volgorde
            diff_df = diff_df[['Omschrijving', 'Vorig aantal', 'Nieuw aantal', 'Verschil']]
            
            if not diff_df.empty:
                st.markdown("### Overzicht van verschillen")
                # Functie voor het kleuren van de 'Verschil'-kolom
                def color_diff(val):
                    try:
                        if val > 0:
                            return 'color: green; font-weight: bold'
                        elif val < 0:
                            return 'color: red; font-weight: bold'
                        else:
                            return ''
                    except:
                        return ''
                styled_diff = diff_df.style.applymap(color_diff, subset=['Verschil'])
                st.markdown(styled_diff.to_html(), unsafe_allow_html=True)
            else:
                st.info("Geen verschillen gevonden tussen de catalogus en de stocklijst.")
        except Exception as e:
            st.error(f"Fout bij het genereren van het overzicht van verschillen: {e}")
st.markdown('</div>', unsafe_allow_html=True)
