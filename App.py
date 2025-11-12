import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Fraksjonsoversikt", layout="wide")
st.title("Fraksjonsoversikt")

uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx"])

if uploaded_file:
    # Les Excel (krever openpyxl i requirements)
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mangler kolonner: {missing}")
        st.stop()

    # Fraksjonsnavn iht. regelen din
    df["Fraksjon"] = df.apply(
        lambda x: x["Materialkorttekst"] if str(x["Betegnelse"]) == "Kranbil Isekk - Avfallstype" else x["Betegnelse"],
        axis=1
    )

    # Rydd tall og enheter
    df["Målkvantum"] = pd.to_numeric(df["Målkvantum"], errors="coerce").fillna(0)
    df["KE.1"] = df["KE.1"].astype(str).str.upper().str.strip()

    # Enheter som finnes i data (f.eks. KG, ST)
    units_found = sorted(df["KE.1"].dropna().unique().tolist())
    # Vis KG-kolonne først, deretter de andre i alfabetisk rekkefølge
    units_order = (["KG"] if "KG" in units_found else []) + [u for u in units_found if u != "KG"]

    # Pivot uten fill_value -> gir NaN der det ikke finnes verdi
    pivot_raw = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum"
    ).reindex(columns=units_order)

    # Total (NaN behandles som 0 i summen)
    totals_col = pivot_raw.fillna(0).sum(axis=1)
    pivot = pivot_raw.copy()
    pivot["Total"] = totals_col

    # --- RAD-REKKEFØLGE ---
    # group_priority: 0 = har KG, 1 = har ikke KG men har ST, 2 = ingen av delene
    kg_vals = pivot_raw["KG"].fillna(0) if "KG" in pivot_raw.columns else pd.Series(0, index=pivot_raw.index)
    st_vals = pivot_raw["ST"].fillna(0) if "ST" in pivot_raw.columns else pd.Series(0, index=pivot_raw.index)
    group_priority = (
        0 * (kg_vals > 0) +
        1 * ((kg_vals <= 0) & (st_vals > 0)) +
        2 * ((kg_vals <= 0) & (st_vals <= 0))
    )

    pivot = pivot.assign(_grp=group_priority).sort_values(by=["_grp", "Total"], ascending=[True, False]).drop(columns="_grp")

    # SUM-rad
    grand_totals = pivot.fillna(0).sum(axis=0).to_frame().T
    grand_totals.index = ["SUM"]
    result = pd.concat([pivot, grand_totals], axis=0)

    # Visning: tomme celler der det ikke finnes verdi
    display_df = result.copy()
    for col in units_order:
        display_df[col] = display_df[col].apply(lambda v: "" if pd.isna(v) else int(v) if float(v).is_integer() else v)
    display_df["Total"] = display_df["Total"].apply(lambda v: int(v) if float(v).is_integer() else v)

    st.subheader("Resultat")
    st.dataframe(display_df, use_container_width=True)

    # Nedlasting til Excel (behold tomme celler i enhetskolonnene)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df = result.copy()
        for col in units_order:
            export_df[col] = export_df[col].where(~export_df[col].isna(), "")
        export_df.to_excel(writer, sheet_name="Fraksjonsoversikt")
    output.seek(0)

    st.download_button(
        label="Last ned som Excel",
        data=output.getvalue(),
        file_name="fraksjonsoversikt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
        # --- Personvern ---
    st.caption("Personvern: Opplastede filer behandles i minnet i din økt og lagres ikke.")
else:
    st.info("Last opp en Excel-fil for å starte.")
