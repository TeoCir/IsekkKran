import pandas as pd
import streamlit as st
from io import BytesIO

st.title("Fraksjonsoversikt")

# Last opp Excel-fil
uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx"])

if uploaded_file:
    # Les data
    df = pd.read_excel(uploaded_file)

    # Sjekk at nødvendige kolonner finnes
    required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1"]
    if all(col in df.columns for col in required_cols):
        # Lag hjelpekolonne for Fraksjon
        df["Fraksjon"] = df.apply(
            lambda x: x["Materialkorttekst"] if x["Betegnelse"] == "Kranbil Isekk - Avfallstype" else x["Betegnelse"],
            axis=1
        )

        # Pivot-tabell
        pivot = df.pivot_table(
            index="Fraksjon",
            columns="KE.1",
            values="Målkvantum",
            aggfunc="sum",
            fill_value=0,
            margins=True,
            margins_name="Total"
        )

        st.subheader("Resultat")
        st.dataframe(pivot)

        # Last ned som Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pivot.to_excel(writer, sheet_name="Fraksjonsoversikt")
        st.download_button(
            label="Last ned som Excel",
            data=output.getvalue(),
            file_name="fraksjonsoversikt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error(f"Mangler en eller flere kolonner: {required_cols}")
