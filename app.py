import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import load_workbook

def read_protected_excel(file):
    """
    Probeer een beveiligd Excel-bestand in te lezen en negeer stijlinformatie.
    """
    try:
        # Probeer direct in te lezen zonder stijlen
        df = pd.read_excel(file, engine="openpyxl", skiprows=0)
        return df
    except Exception as e:
        try:
            # Als direct lezen mislukt, probeer handmatig laden zonder stijlen
            wb = load_workbook(file, read_only=True, data_only=True)
            sheet = wb.active

            # Lees de data handmatig in
            data = sheet.values
            columns = next(data)  # Haal de kolomnamen uit de eerste rij
            df = pd.DataFrame(data, columns=columns)
            return df
        except Exception as inner_e:
            st.error("Fout bij het inlezen van de stocklijst: Mogelijk bevat het bestand corrupte of complexe stijlen.")
            st.error(f"Technische foutmelding: {inner_e}")
            return pd.DataFrame()

def filter_stock(stock_file, catalog_file):
    try:
        # Stocklijst inlezen
        stocklijst_df = read_protected_excel(stock_file)
        if stocklijst_df.empty:
            st.error("De stocklijst is leeg of kan niet worden gelezen. Controleer of het bestand correct is opgeslagen.")
            return pd.DataFrame()
    except Exception as e:
        st.error(f"Fout bij het verwerken van de stocklijst: {e}")
        return pd.DataFrame()

    try:
        # Catalogus inlezen met automatische delimiter detectie
        catalogus_df = pd.read_csv(catalog_file, sep=None, engine="python")
    except Exception as e:
        st.error(f"Fout bij het inlezen van de catalogus: {e}")
        return pd.DataFrame()

    # Kolommen identificeren voor filtering
    stocklijst_col = "Code"  # Alternatief: "EAN"
    catalogus_col = "product_sku"
    catalogus_name_col = "product_name"

    try:
        # Converteren naar string om mogelijke datatypeverschillen te voorkomen
        stocklijst_df[stocklijst_col] = stocklijst_df[stocklijst_col].astype(str)
        catalogus_df[catalogus_col] = catalogus_df[catalogus_col].astype(str)
    except KeyError as e:
        st.error(f"Vereiste kolom ontbreekt in de bestanden: {e}")
        return pd.DataFrame()

    try:
        # Filteren: Alleen rijen uit de stocklijst behouden die in de catalogus staan
        filtered_stocklijst_df = stocklijst_df[stocklijst_df[stocklijst_col].isin(catalogus_df[catalogus_col])]

        # Toevoegen van product_name vanuit de catalogus
        merged_df = filtered_stocklijst_df.merge(
            catalogus_df[[catalogus_col, catalogus_name_col]],
            left_on=stocklijst_col,
            right_on=catalogus_col,
            how="left"
        )
    except Exception as e:
        st.error(f"Fout bij het filteren en samenvoegen van data: {e}")
        return pd.DataFrame()

    # Kolommen hernoemen en filteren voor export
    rename_map = {
        "Code": "product_sku",
        "# stock": "product_quantity"
    }

    try:
        # Controleer of alle vereiste kolommen beschikbaar zijn
        missing_columns = [col for col in rename_map.keys() if col not in merged_df.columns]
        if missing_columns:
            st.error(f"De volgende vereiste kolommen ontbreken in de stocklijst: {missing_columns}")
            return pd.DataFrame()

        merged_df = merged_df.rename(columns=rename_map)

        # Alleen gewenste kolommen behouden
        merged_df = merged_df[[
            "product_name", "product_sku", "product_quantity"
        ]]

        # product_quantity omzetten naar gehele getallen
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
