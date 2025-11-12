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

    # Fjern "NAN"/blanke enheter
    df = df[df["KE.1"].ne("NAN") & df["KE.1"].ne("")]

    # Enheter og rekkefølge (KG først)
    units_found = sorted(df["KE.1"].dropna().unique().tolist())
    units_order = (["KG"] if "KG" in units_found else []) + [u for u in units_found if u != "KG"]

    # Pivot (tomt der det ikke finnes verdi)
    pivot = df.pivot_table(
        index="Fraksjon",
        columns="KE.1",
        values="Målkvantum",
        aggfunc="sum"
    ).reindex(columns=units_order)

    # SUM-rad (men ingen Total-kolonne)
    sum_row = pivot.fillna(0).sum(axis=0).to_frame().T
    sum_row.index = ["SUM"]
    result = pd.concat([pivot, sum_row], axis=0)

    # Sortér fraksjoner (før SUM): KG først (mest KG), så ST (mest ST)
    data_part = result.loc[result.index != "SUM"]
    kg_vals = data_part["KG"].fillna(0) if "KG" in data_part.columns else pd.Series(0, index=data_part.index)
    st_vals = data_part["ST"].fillna(0) if "ST" in data_part.columns else pd.Series(0, index=data_part.index)
    group_priority = (
        0 * (kg_vals > 0)
        + 1 * ((kg_vals <= 0) & (st_vals > 0))
        + 2 * ((kg_vals <= 0) & (st_vals <= 0))
    )
    sort_df = pd.DataFrame({"_grp": group_priority, "_kg": kg_vals, "_st": st_vals}, index=data_part.index)
    ordered_index = sort_df.sort_values(by=["_grp", "_kg", "_st"], ascending=[True, False, False]).index
    data_part = data_part.loc[ordered_index]
    # Slå sammen igjen med SUM nederst
    if "SUM" in result.index:
        result = pd.concat([data_part, result.loc[["SUM"]]], axis=0)
    else:
        result = data_part

    # ---- VISNINGSBRYTER ----
    view = st.radio(
        "Visning",
        ["Standard (pivot)", "Pen visning (chips)"],
        horizontal=True,
        index=0
    )

    if view == "Standard (pivot)":
        # Visning: tomme celler (ikke 0)
        display_df = result.copy()
        for col in units_order:
            display_df[col] = display_df[col].apply(lambda v: "" if pd.isna(v) else int(v) if float(v).is_integer() else v)
        st.subheader("Resultat")
        st.dataframe(display_df, use_container_width=True)

    else:
        # Pen visning: kun enheter som finnes per fraksjon som "chips"
        st.subheader("Resultat (pen visning)")
        # CSS for chips
        st.markdown("""
        <style>
        .chip {display:inline-block; padding:4px 8px; margin:2px 6px 2px 0;
               border-radius:999px; background:#eef1f5; font-size:0.9rem;}
        .row {display:flex; justify-content:space-between; align-items:center;
              padding:8px 0; border-bottom:1px solid #eee;}
        .name {font-weight:600;}
        .sumrow {background:#fafafa; border-top:2px solid #ddd; padding:10px 0; margin-top:6px;}
        </style>
        """, unsafe_allow_html=True)

        # Bygg chips per rad
        def chips_for_row(s):
            items = []
            for u in units_order:
                val = s.get(u, float("nan"))
                if pd.notna(val) and float(val) != 0.0:
                    shown = int(val) if float(val).is_integer() else round(float(val), 3)
                    items.append(f"<span class='chip'>{shown} {u}</span>")
            return "".join(items)

        # Tegn rader (uten SUM)
        for idx, row in result.loc[result.index != "SUM"].iterrows():
            chips_html = chips_for_row(row)
            st.markdown(
                f"<div class='row'><div class='name'>{idx}</div><div>{chips_html}</div></div>",
                unsafe_allow_html=True
            )

        # SUM-rad som chips
        if "SUM" in result.index:
            sum_chips = chips_for_row(result.loc["SUM"])
            st.markdown(f"<div class='row sumrow'><div class='name'>SUM</div><div>{sum_chips}</div></div>", unsafe_allow_html=True)

    # ---- Nedlasting til Excel (pivot-format) ----
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

    # Personvern
    st.caption("Personvern: Opplastede filer behandles i minnet i din økt og lagres ikke.")
else:
    st.info("Last opp en Excel-fil for å starte.")
