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

# --- Functie om Excel in te lezen ---
def read_excel_simple(file):
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

# --- Functie voor het filteren van de data ---
def filter_stock(stock_file, catalog_file, progress_callback=None):
    if progress_callback:
        progress_callback(10)
    stocklijst_df = read_excel_simple(stock_file)
    if stocklijst_df.empty:
        st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
        return pd.DataFrame()

    if progress_callback:
        progress_callback(30)
    try:
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()

    stocklijst_col = "Code"      # Bijvoorbeeld: "Code" (of "EAN")
    catalogus_col = "product_sku"
    
    if progress_callback:
        progress_callback(50)
    try:
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()

    try:
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]
    except Exception as e:
        st.error(f"Fout bij het filteren van data: {e}")
        return pd.DataFrame()

    if progress_callback:
        progress_callback(70)
    try:
        merged_df = filtered_stocklijst_df.merge(
            catalogus_df[[catalogus_col]],
            left_on=stocklijst_col,
            right_on=catalogus_col,
            how="left"
        )
        merged_df = merged_df.drop(columns=[catalogus_col])
    except Exception as e:
        st.error(f"Fout bij het samenvoegen van data: {e}")
        return pd.DataFrame()

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
        merged_df["product_quantity"] = pd.to_numeric(merged_df["product_quantity"], errors="coerce").fillna(0).astype(int)
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
            # Reset de pointer van catalog_file zodat deze opnieuw ingelezen kan worden
            catalog_file.seek(0)
            try:
                catalogus_df_full = pd.read_csv(catalog_file, sep=None, engine="python")
            except Exception as e:
                try:
                    catalogus_df_full = pd.read_csv(catalog_file, delimiter=',')
                except Exception as e:
                    catalogus_df_full = pd.read_csv(catalog_file, delimiter=';')
            
            catalogus_df_full['product_sku'] = catalogus_df_full['product_sku'].astype(str)
            # Gebruik "product_name" als de productnaam; indien niet aanwezig, standaardwaarde "Geen productnaam"
            if 'product_name' not in catalogus_df_full.columns:
                catalogus_df_full['product_name'] = "Geen productnaam"
            # We nemen ook de kolom product_price mee voor de omzetberekening
            if 'product_price' not in catalogus_df_full.columns:
                st.error("De kolom 'product_price' ontbreekt in de catalogus.")
            else:
                # Zorg ervoor dat product_price numeriek is
                catalogus_df_full['product_price'] = pd.to_numeric(catalogus_df_full['product_price'], errors='coerce').fillna(0)

            # Maak een kopie van de export en hernoem de stocklijsthoeveelheid naar 'Nieuw aantal'
            filtered_export = filtered_df.copy().rename(columns={'product_quantity': 'Nieuw aantal'})
            
            # Voeg de catalogus-informatie toe: gebruik "product_name" als "Productnaam", 
            # "product_quantity" als "Vorig aantal" en neem "product_price" mee.
            diff_df = pd.merge(filtered_export, 
                               catalogus_df_full[['product_sku', 'product_name', 'product_quantity', 'product_price']], 
                               on='product_sku', how='left')
            diff_df = diff_df.rename(columns={'product_quantity': 'Vorig aantal', 
                                              'product_name': 'Productnaam'})
            # Zorg ervoor dat 'Vorig aantal' een geheel getal is
            diff_df['Vorig aantal'] = pd.to_numeric(diff_df['Vorig aantal'], errors='coerce').fillna(0).astype(int)
            
            # Bereken het verschil: (Nieuw aantal - Vorig aantal)
            diff_df['Verschil'] = diff_df['Nieuw aantal'] - diff_df['Vorig aantal']
            # Bereken de omzet: voor negatieve verschillen (verloren voorraad) is omzet = (-Verschil) * product_price
            diff_df['omzet'] = diff_df.apply(lambda row: (-row['Verschil'] * row['product_price']) if row['Verschil'] < 0 else 0, axis=1)
            # Houd enkel producten met een verschil
            diff_df = diff_df[diff_df['Verschil'] != 0]
            # Behoud de gewenste kolommen in de volgorde
            diff_df = diff_df[['Productnaam', 'Vorig aantal', 'Nieuw aantal', 'Verschil', 'omzet']]
            
            if not diff_df.empty:
                # --- Paginering ---
                # Als er meer dan 30 rijen zijn, splits het overzicht op in pagina's van 30
                if 'diff_page' not in st.session_state:
                    st.session_state.diff_page = 0
                per_page = 30
                total_rows = len(diff_df)
                total_pages = (total_rows - 1) // per_page + 1
                start_idx = st.session_state.diff_page * per_page
                end_idx = start_idx + per_page
                diff_page = diff_df.iloc[start_idx:end_idx]

                # Functie voor rij-opmaak: kleur de hele rij afhankelijk van het verschil
                def row_style(row):
                    if row['Verschil'] < 0:
                        return ['background-color: lightgreen'] * len(row)
                    elif row['Verschil'] > 0:
                        return ['background-color: lightcoral'] * len(row)
                    else:
                        return [''] * len(row)
                
                styled_diff = diff_page.style.apply(row_style, axis=1)
                # Centreer de gehele tabel met behulp van CSS
                table_html = styled_diff.to_html()
                st.markdown(f"<div style='display: flex; justify-content: center;'>{table_html}</div>", unsafe_allow_html=True)
                
                # Toon een centrale titel voor het overzicht
                st.markdown("<h3 style='text-align: center;'>Overzicht van verschillen</h3>", unsafe_allow_html=True)
                
                # Navigatieknoppen voor paginering, gecentreerd
                nav_cols = st.columns(3)
                if st.session_state.diff_page > 0:
                    if nav_cols[0].button("Vorige 30"):
                        st.session_state.diff_page -= 1
                        st.experimental_rerun()
                else:
                    nav_cols[0].empty()
                nav_cols[1].markdown(f"<div style='text-align: center;'>Pagina {st.session_state.diff_page + 1} van {total_pages}</div>", unsafe_allow_html=True)
                if end_idx < total_rows:
                    if nav_cols[2].button("Volgende 30"):
                        st.session_state.diff_page += 1
                        st.experimental_rerun()
                else:
                    nav_cols[2].empty()
            else:
                st.info("Geen verschillen gevonden tussen de catalogus en de stocklijst.")
        except Exception as e:
            st.error(f"Fout bij het genereren van het overzicht van verschillen: {e}")
st.markdown('</div>', unsafe_allow_html=True)
