import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Fraksjonsoversikt", layout="wide")
st.title("Fraksjonsoversikt")

uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx"])

BAD_UNIT_LABELS = {"", "NAN", "NA", "NONE", "NULL", "TOTAL", "SUM"}

def clean_unit(u):
    """Rens enhet (KE.1) og fjern søppelverdier."""
    if pd.isna(u):  # ekte NaN
        return None
    s = str(u).strip().upper()
    return None if s in BAD_UNIT_LABELS else s

def fmt_number(val):
    """Formater tall for visning i tabell (blank hvis NaN)."""
    if pd.isna(val):
        return ""
    f = float(val)
    if f.is_integer():
        return str(int(f))
    return str(f)

if uploaded_file:
    # Les Excel (krever openpyxl i requirements.txt)
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mangler kolonner: {missing}")
        st.stop()

    # Fraksjon etter regelen din
    df["Fraksjon"] = df.apply(
        lambda x: x["Materialkorttekst"]
        if str(x["Betegnelse"]) == "Kranbil Isekk - Avfallstype"
        else x["Betegnelse"],
        axis=1,
    )

    # Rens tall + enhet
    # Viktig: ikke fillna(0) her – vi vil beholde "ingen verdi" som NaN
    df["Målkvantum"] = pd.to_numeric(df["Målkvantum"], errors="coerce")
    df["KE.1"] = df["KE.1"].map(clean_unit)
    df = df[df["KE.1"].notna()].copy()

    # Enhetsrekkefølge (KG først, deretter alfabetisk på resten)
    units_found = sorted(df["KE.1"].unique().tolist())
    units_order = (["KG"] if "KG" in units_found else []) + [
        u for u in units_found if u != "KG"
    ]

    # Pivot 1: summerte verdier
    pivot_vals = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum",
    )

    # Pivot 2: sjekk om det fantes noen faktiske verdier (ikke bare tomt)
    pivot_has = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc=lambda s: s.notna().any(),
    )

    # Ekstra sikring mot rare kolonner
    safe_cols = [
        c
        for c in pivot_vals.columns
        if (pd.notna(c) and str(c).strip() and str(c).strip().upper() not in BAD_UNIT_LABELS)
    ]
    pivot_vals = pivot_vals.loc[:, safe_cols]
    pivot_has = pivot_has.loc[:, safe_cols]

    # Rekkefølge på enheter
    pivot_vals = pivot_vals.reindex(columns=units_order)
    pivot_has = pivot_has.reindex(columns=units_order)

    # Hvis en fraksjon/enhet bare hadde tomme verdier → sett til NaN (for blank visning)
    pivot = pivot_vals.where(pivot_has, pd.NA)

    # SUM-rad (summerer per enhet, tomme behandles som 0 i summen)
    sum_row = pivot.fillna(0).sum(axis=0).to_frame().T
    sum_row.index = ["SUM"]
    result = pd.concat([pivot, sum_row], axis=0)

    # Sortering av fraksjoner (før SUM): KG først (mest KG), så ST (mest ST), så resten
    data = result.loc[result.index != "SUM"]
    kg_vals = data["KG"].fillna(0) if "KG" in data.columns else pd.Series(0, index=data.index)
    st_vals = data["ST"].fillna(0) if "ST" in data.columns else pd.Series(0, index=data.index)
    grp = (
        0 * (kg_vals > 0)
        + 1 * ((kg_vals <= 0) & (st_vals > 0))
        + 2 * ((kg_vals <= 0) & (st_vals <= 0))
    )
    order = (
        pd.DataFrame({"_g": grp, "_kg": kg_vals, "_st": st_vals}, index=data.index)
        .sort_values(by=["_g", "_kg", "_st"], ascending=[True, False, False])
        .index
    )
    result = pd.concat([data.loc[order], result.loc[["SUM"]]], axis=0)

    # ---------- VISNINGSTABELL ----------
    disp = result.copy()

    # Formater tall og blanke felt
    for col in disp.columns:
        disp[col] = disp[col].apply(fmt_number)

    # Flytt index til kolonne "Fraksjon"
    disp = disp.reset_index()
    if disp.columns[0] != "Fraksjon":
        disp = disp.rename(columns={disp.columns[0]: "Fraksjon"})

    # Litt CSS for å få tabellen bred
    st.markdown(
        """
        <style>
        table {
            width: 100% !important;
        }
        th, td {
            padding: 4px 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Style Fraksjon-kolonnen fet og litt større
    styled = (
        disp.style
        .set_properties(subset=["Fraksjon"], **{"font-weight": "bold", "color": "black", "font-size": "15px"})
        .set_properties(**{"font-size": "14px"})
        .set_table_styles(
            [
                {"selector": "table", "props": [("width", "100%")]},
                {"selector": "th", "props": [("text-align", "left")]}
            ]
        )
    )

    st.subheader("Resultat")
    st.write(styled.to_html(), unsafe_allow_html=True)

    # ---------- EXCEL-NEDLASTING ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:
        export_df = result.copy()
        for col in export_df.columns:
            export_df[col] = export_df[col].where(~export_df[col].isna(), "")
        export_df.to_excel(w, sheet_name="Fraksjonsoversikt", index=True)
    output.seek(0)
    st.download_button(
        "Last ned som Excel",
        output.getvalue(),
        file_name="fraksjonsoversikt.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Personvern
    st.caption("Personvern: Opplastede filer behandles i minnet i din økt og lagres ikke.")
else:
    st.info("Last opp en Excel-fil for å starte.")
