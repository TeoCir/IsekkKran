import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Fraksjonsoversikt", layout="wide")
st.title("Fraksjonsoversikt")

uploaded_file = st.file_uploader("Last opp Excel-fil", type=["xlsx"])

BAD_UNIT_LABELS = {"", "NAN", "NA", "NONE", "NULL", "TOTAL", "SUM"}

def clean_unit(u):
    if pd.isna(u):  # ekte NaN
        return None
    s = str(u).strip().upper()
    return None if s in BAD_UNIT_LABELS else s

def fmt_number(val, decimals=0):
    if pd.isna(val):
        return ""
    f = float(val)
    if f.is_integer():
        return str(int(f))
    return f"{f:.{decimals}f}"

if uploaded_file:
    # Les Excel
    df = pd.read_excel(uploaded_file, engine="openpyxl")

    required_cols = ["Betegnelse", "Materialkorttekst", "Målkvantum", "KE.1"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.error(f"Mangler kolonner: {missing}")
        st.stop()

    # Fraksjon etter regelen din
    df["Fraksjon"] = df.apply(
        lambda x: x["Materialkorttekst"] if str(x["Betegnelse"]) == "Kranbil Isekk - Avfallstype" else x["Betegnelse"],
        axis=1
    )

    # Rens tall + enhet
    df["Målkvantum"] = pd.to_numeric(df["Målkvantum"], errors="coerce").fillna(0)
    df["KE.1"] = df["KE.1"].map(clean_unit)
    df = df[df["KE.1"].notna()].copy()

    # Enhetsrekkefølge (KG først)
    units_found = sorted(df["KE.1"].unique().tolist())
    units_order = (["KG"] if "KG" in units_found else []) + [u for u in units_found if u != "KG"]

    # Pivot (tomme celler der det ikke finnes verdi)
    pivot = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum"
    )

    # Ekstra sikring mot rare kolonner
    safe_cols = [
        c for c in pivot.columns
        if (pd.notna(c) and str(c).strip() and str(c).strip().upper() not in BAD_UNIT_LABELS)
    ]
    pivot = pivot.loc[:, safe_cols].reindex(columns=units_order)

    # SUM-rad
    sum_row = pivot.fillna(0).sum(axis=0).to_frame().T
    sum_row.index = ["SUM"]
    result = pd.concat([pivot, sum_row], axis=0)

    # Sortering (før SUM): KG først (mest KG), så ST (mest ST), så resten
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
    # Blank i UI ved NaN
    for col in disp.columns:
        disp[col] = disp[col].apply(
            lambda v: "" if pd.isna(v) else int(v) if float(v).is_integer() else v
        )

    # Reset index for å få Fraksjon som kolonne
    disp = disp.reset_index()  # index-navnet er "Fraksjon"
    # Sørg for at kolonnen heter Fraksjon
    if disp.columns[0] != "Fraksjon":
        disp = disp.rename(columns={disp.columns[0]: "Fraksjon"})

    styled = (
        disp.style
        .set_properties(subset=["Fraksjon"], **{"font-weight": "bold", "color": "black", "font-size": "15px"})
        .set_properties(**{"font-size": "14px"})
    )

    st.subheader("Resultat")
    st.write(styled.to_html(), unsafe_allow_html=True)

    # ---------- PUNCH-LINJER ----------
    st.subheader("Punch-linjer (klar for kopiering)")
    col1, col2, col3 = st.columns([1.2, 1, 1])
    with col1:
        sep_name = st.selectbox("Separator", ["Tab (anbefalt)", ";", ",", "|"], index=0)
    with col2:
        include_sum = st.checkbox("Ta med SUM", value=False)
    with col3:
        decimals = st.number_input("Desimaler ved behov", min_value=0, max_value=6, value=0, step=1)

    sep_map = {
        "Tab (anbefalt)": "\t",
        ";": ";",
        ",": ",",
        "|": "|"
    }
    sep = sep_map[sep_name]

    # Bygg punch-linjer: Fraksjon + én kolonne per enhet i fast rekkefølge
    punch_df = result.copy()
    if not include_sum and "SUM" in punch_df.index:
        punch_df = punch_df.loc[punch_df.index != "SUM"]

    punch_df = punch_df.reindex(columns=units_order)

    out_rows = []
    header = ["Fraksjon"] + units_order
    out_rows.append(sep.join(header))

    for idx, row in punch_df.iterrows():
        # hent råverdier for enhetene
        raw_vals = [row.get(u) for u in units_order]
        # sjekk om ALT er tomt/NaN
        if all(pd.isna(v) for v in raw_vals):
            # hele fraksjonen er "tom" -> ikke sett 0, bare blanke felter
            vals = ["" for _ in units_order]
        else:
            # minst én enhet har verdi -> manglende enheter blir 0
            vals = []
            for v in raw_vals:
                if pd.isna(v):
                    v = 0
                vals.append(fmt_number(v, decimals=decimals))

        out_rows.append(sep.join([str(idx)] + vals))

    punch_text = "\n".join(out_rows)

    # Kopier-knapp via st.code
    st.code(punch_text, language="text")

    # Last ned punch-linjer
    mime = "text/plain" if sep != ";" else "text/csv"
    st.download_button(
        "Last ned punch-linjer",
        punch_text.encode("utf-8"),
        file_name="punch-linjer.txt" if mime == "text/plain" else "punch-linjer.csv",
        mime=mime
    )

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
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Personvern
    st.caption("Personvern: Opplastede filer behandles i minnet i din økt og lagres ikke.")
else:
    st.info("Last opp en Excel-fil for å starte.")
