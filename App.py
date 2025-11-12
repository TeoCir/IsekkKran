import pandas as pd
import streamlit as st
from io import BytesIO
from pathlib import Path

st.set_page_config(page_title="Fraksjonsoversikt", layout="wide")
st.title("Fraksjonsoversikt")

# --- Opplasting ---
uploaded_file = st.file_uploader("Last opp fil", type=["xlsx", "xls", "csv"])

def read_any(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, sep=None, engine="python")
    if name.endswith(".xlsx"):
        # Viktig: krever openpyxl i requirements
        return pd.read_excel(uploaded, engine="openpyxl")
    if name.endswith(".xls"):
        # Krever xlrd==1.2.0 hvis du vil støtte gamle .xls
        return pd.read_excel(uploaded, engine="xlrd")
    raise ValueError("Ukjent filtype")

if uploaded_file:
    try:
        df = read_any(uploaded_file)
    except Exception as e:
        st.error(f"Kunne ikke lese filen: {e}")
        st.stop()

    # --- Sjekk kolonner ---
    required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mangler kolonner: {missing}")
        st.stop()

    # --- Valg av enheter (default KG) ---
    units = sorted(df["KE.1"].dropna().astype(str).unique().tolist())
    default_units = ["KG"] if "KG" in units else units
    chosen_units = st.multiselect("Hvilke enheter skal telle med?", units, default=default_units)

    if not chosen_units:
        st.warning("Velg minst én enhet.")
        st.stop()

    # --- Lag fraksjonsnavn iht. regelen din ---
    df["Fraksjon"] = df.apply(
        lambda x: x["Materialkorttekst"] if str(x["Betegnelse"]) == "Kranbil Isekk - Avfallstype" else x["Betegnelse"],
        axis=1
    )

    # --- Filtrer på enheter og rydd tall ---
    work = df[df["KE.1"].astype(str).isin(chosen_units)].copy()
    work["Målkvantum"] = pd.to_numeric(work["Målkvantum"], errors="coerce").fillna(0)

    # --- Pivot ---
    pivot = work.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum",
        fill_value=0,
    )

    # Total-kolonne på tvers av valgte enheter
    pivot["Total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("Total", ascending=False)

    # Total-rad
    totals_row = pd.DataFrame(pivot.sum(axis=0)).T
    totals_row.index = ["SUM"]
    pivot_with_sum = pd.concat([pivot, totals_row], axis=0)

    st.subheader("Resultat")
    st.dataframe(pivot_with_sum, use_container_width=True)

    # --- Nedlasting til Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pivot_with_sum.to_excel(writer, sheet_name="Fraksjonsoversikt")
    st.download_button(
        label="Last ned som Excel",
        data=output.getvalue(),
        file_name="fraksjonsoversikt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Last opp en Excel/CSV for å starte.")

