import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Fraksjonsoversikt", layout="wide")
st.title("Fraksjonsoversikt")

uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx"])

def _clean_unit(u):
    if pd.isna(u):
        return None
    s = str(u).strip().upper()
    # fjern søppelverdier som lager uønskede kolonner
    if s in {"", "NAN", "NA", "NONE", "NULL", "TOTAL", "SUM"}:
        return None
    return s

if uploaded_file:
    # Les Excel (krever openpyxl i requirements.txt)
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
    df["KE.1"] = df["KE.1"].map(_clean_unit)
    df = df[df["KE.1"].notna()].copy()

    # Enheter og rekkefølge: KG først
    units_found = sorted(df["KE.1"].unique().tolist())
    units_order = (["KG"] if "KG" in units_found else []) + [u for u in units_found if u != "KG"]

    # Pivot (tomme celler der det ikke finnes verdi)
    pivot = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum"
    )

    # Ekstra sikring: dropp eventuelle rare kolonner
    safe_cols = [c for c in pivot.columns if pd.notna(c) and str(c).strip() and str(c).strip().upper() not in {"TOTAL", "SUM", "NAN"}]
    pivot = pivot.loc[:, safe_cols].reindex(columns=units_order)

    # SUM-rad (ingen Total-kolonne)
    sum_row = pivot.fillna(0).sum(axis=0).to_frame().T
    sum_row.index = ["SUM"]
    result = pd.concat([pivot, sum_row], axis=0)

    # Sorter fraksjoner (før SUM): KG først (mest KG), så ST (mest ST), så resten
    data_part = result.loc[result.index != "SUM"]
    kg_vals = data_part["KG"].fillna(0) if "KG" in data_part.columns else pd.Series(0, index=data_part.index)
    st_vals = data_part["ST"].fillna(0) if "ST" in data_part.columns else pd.Series(0, index=data_part.index)
    group_priority = (
        0 * (kg_vals > 0)
        + 1 * ((kg_vals <= 0) & (st_vals > 0))
        + 2 * ((kg_vals <= 0) & (st_vals <= 0))
    )
    order_idx = (
        pd.DataFrame({"_grp": group_priority, "_kg": kg_vals, "_st": st_vals}, index=data_part.index)
        .sort_values(by=["_grp", "_kg", "_st"], ascending=[True, False, False])
        .index
    )
    data_part = data_part.loc[order_idx]

    # Sett sammen igjen med SUM nederst
    result = pd.concat([data_part, result.loc[["SUM"]]], axis=0)

    # Visning: tomme celler (ikke 0)
    display_df = result.copy()
    for col in display_df.columns:
        display_df[col] = display_df[col].apply(lambda v: "" if pd.isna(v) else int(v) if float(v).is_integer() else v)

    st.subheader("Resultat")
    st.dataframe(display_df, use_container_width=True)

    # Nedlasting til Excel (pivot + SUM, uten Total/NaN-kolonner)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df = result.copy()
        for col in export_df.columns:
            export_df[col] = export_df[col].where(~export_df[col].isna(), "")
        export_df.to_excel(writer, sheet_name="Fraksjonsoversikt")
    output.seek(0)

    st.download_button(
        label="Last ned som Excel",
        data=output.getvalue(),
        file_name="fraksjonsoversikt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.caption("Personvern: Opplastede filer behandles i minnet i din økt og lagres ikke.")
else:
    st.info("Last opp en Excel-fil for å starte.")
