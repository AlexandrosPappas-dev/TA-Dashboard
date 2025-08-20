import streamlit as st
import pandas as pd
import os
import glob
import altair as alt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image as RLImage, Table, TableStyle, PageBreak
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
import matplotlib.pyplot as plt
import io
import base64
from altair_saver import save
import vl_convert as vlc

# === Psychography colors ===
PSYCHOGRAPHIE_FARBEN = {
    "AE": "rgb(104, 125, 1)",
    "AH": "rgb(167, 202, 2)",
    "AR": "rgb(217, 255, 28)",
    "LE": "rgb(35, 67, 84)",
    "LH": "rgb(85, 134, 161)",
    "LR": "rgb(147, 206, 237)",
    "ME": "rgb(178, 34, 34)",
    "MH": "rgb(239, 44, 53)",
    "MR": "rgb(255, 71, 81)",
    "all": "rgb(81, 71, 31)"
}

# === Customer Journey Steps ===
JOURNEY_STEPS = ["Awareness", "Consideration", "Purchase", "Satisfaction", "Loyalty"]

st.set_page_config(page_title="TA Dashboard", layout="wide")
st.title("Driver Analysis Dashboard – Cadillac x Cronbach 2025")

# === Load data ===
@st.cache_data
def lade_daten(datengruppen=["Detail", "Cluster"], basisordner="data"):
    datensaetze = []
    for projekt_ordner in glob.glob(os.path.join(basisordner, "*")):
        if not os.path.isdir(projekt_ordner):
            continue
        projekt_name = os.path.basename(projekt_ordner)

        for ebene_ordner in glob.glob(os.path.join(projekt_ordner, "*")):
            if not os.path.isdir(ebene_ordner):
                continue
            ordnerebene = os.path.basename(ebene_ordner)
            if ordnerebene not in ["GFactor", "Results"]:
                continue

            for land_ordner in glob.glob(os.path.join(ebene_ordner, "*")):
                if not os.path.isdir(land_ordner):
                    continue
                land_code = os.path.basename(land_ordner).split("_")[-1]

                for psychographie_ordner in glob.glob(os.path.join(land_ordner, "*")):
                    if not os.path.isdir(psychographie_ordner):
                        continue
                    psychographie = os.path.basename(psychographie_ordner)

                    for datengruppe in datengruppen:
                        datengruppe_ordner = os.path.join(psychographie_ordner, datengruppe)
                        if not os.path.isdir(datengruppe_ordner):
                            continue

                        for excel_datei in glob.glob(os.path.join(datengruppe_ordner, "*.xlsx")):
                            try:
                                df = pd.read_excel(excel_datei, sheet_name=0, header=None)
                                stage = df.iloc[4, 0]
                                markt = df.iloc[5, 0]
                                treiberset = df.iloc[6, 0]
                                cluster_info = df.iloc[7, 0] if datengruppe == "Cluster" else None

                                r_squared_row = df[df[0] == "R SQUARED MODEL"].index[0]
                                adjusted_r2 = round(float(df.iloc[r_squared_row + 1, 0]), 2)

                                n_row = df[df[0] == "N"].index[0]
                                n_value = int(df.iloc[n_row + 1, 0])

                                dropped_drivers_idx = df[df[0] == '/// DROPPED SIGNIFICANT DRIVERS ///'].index[0] + 1
                                dropped_drivers = df.iloc[dropped_drivers_idx:, 0].dropna().tolist()

                                chart_data = df.iloc[8:, [0, 1]].dropna()
                                chart_data.columns = ["Entity", "Value"]
                                stop_idx = chart_data[chart_data["Entity"] == "Stop"].index
                                if not stop_idx.empty:
                                    chart_data = chart_data.loc[:stop_idx[0] - 1]

                                for _, row in chart_data.iterrows():
                                    eintrag = {
                                        "Project": projekt_name,
                                        "OrdnerEbene": ordnerebene,
                                        "Country": land_code,
                                        "Psychography": psychographie,
                                        "Data Group": datengruppe,
                                        "Stage": stage,
                                        "Market": markt,
                                        "Driver Set": treiberset,
                                        "Entity": row["Entity"],
                                        "Value": row["Value"],
                                        "Adjusted_R2": adjusted_r2,
                                        "n": n_value,
                                        "Dropped_Drivers": dropped_drivers,
                                        "Source": os.path.basename(excel_datei)
                                    }

                                    if cluster_info is not None:
                                        eintrag["Cluster_Info"] = cluster_info

                                    datensaetze.append(eintrag)

                            except Exception as e:
                                st.warning(f"Error in file {excel_datei}: {e}")

    return pd.DataFrame(datensaetze)

df = lade_daten()

# === Filter logic ===
st.sidebar.header("🔎 Filters")

projekt = st.sidebar.selectbox("Select Project", [""] + sorted(df["Project"].unique()))
datengruppe = st.sidebar.selectbox("Select Data Group", [""] + sorted(df["Data Group"].unique()))
general_factor = st.sidebar.selectbox("General Factor Data?", ["No", "Yes"], index=0)

if projekt and datengruppe:
    ordnerebene = "GFactor" if general_factor == "Yes" else "Results"
    df = df[(df["Project"] == projekt) & (df["Data Group"] == datengruppe) & (df["OrdnerEbene"] == ordnerebene)]

    clusterinfo = None
    if datengruppe == "Cluster":
        clusterinfo_opt = sorted(df["Cluster_Info"].dropna().unique())
        clusterinfo = st.sidebar.selectbox("Select Cluster", ["All"] + clusterinfo_opt)
        if clusterinfo != "All":
            df = df[df["Cluster_Info"] == clusterinfo]

    land = st.sidebar.selectbox("Select Country", ["All"] + sorted(df["Country"].unique()))
    if land != "All":
        df = df[df["Country"] == land]

    # Psychography Dropdown mit korrekter Sortierung (All ganz oben, dann alphabetisch)
    psychography_options = ["All"] + sorted([p for p in df["Psychography"].unique() if p != "All"])
    psychographie = st.sidebar.selectbox("Select Psychography", psychography_options)

    if psychographie == "All":
        df = df[df["Psychography"].str.lower() != "all"]
    else:
        df = df[df["Psychography"] == psychographie]
        if psychographie == "all":
            st.warning("This filter includes all psychographies, not only the ones in the drop-down menu. Hence they are based on different data")

    # === Stage Filter NUR WENN general_factor == "No" ===
    if general_factor == "No":
       stage_order = ["Awareness", "Consideration", "Purchase", "Satisfaction", "Loyalty"]
       stage_options = ["All"] + [s for s in stage_order if s in df["Stage"].unique()]
       stage = st.sidebar.selectbox("Select Stage", stage_options)

       if stage != "All":
           df = df[df["Stage"] == stage]
    else:
        stage = None  # Kein Stage-Filter, aber Journey soll angezeigt werden wie bei "All"


    markt = st.sidebar.selectbox("Select Market", ["All"] + sorted(df["Market"].unique()))
    if markt != "All":
        df = df[df["Market"] == markt]

    treiberset_opt = ["Brand", "Product"] if datengruppe == "Cluster" else sorted(df["Driver Set"].dropna().unique())
    treiberset = st.sidebar.selectbox("Select Driver Set", ["All"] + treiberset_opt)
    if treiberset != "All":
        df = df[df["Driver Set"] == treiberset]

    # Visualisierung vorbereiten
    selected_stage = stage if stage and stage != "All" else None
    farbe = PSYCHOGRAPHIE_FARBEN.get(psychographie, "#888") if psychographie != "All" else "#888"


    # =======================
    # === PDF Export Core ===
    # =======================
    # Wichtig: create_pdf_and_download_link akzeptiert jetzt explizit chart_png_bytes,
    # damit IMMER das aktuelle Diagramm in die PDF gelangt (kein session_state nötig).
    def export_to_pdf(filters, chart_df, chart_png_bytes=None, adjusted_r2=None, n_value=None, dropped_drivers=None, values_table_df=None):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer, pagesize=A4,
            leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
        )
        styles = getSampleStyleSheet()
        story = []

        # Header
        title_para = Paragraph("TA Dashboard – Exported View", styles["Title"])
        right_cell = Spacer(1, 1)
        if filters.get("Psychography") not in (None, "", "All"):
            bildpfad = f"graphics/{filters['Psychography']}.png"
            if os.path.exists(bildpfad):
                # Behalte deine bestehende Größe bei (kein unbeauftragter Eingriff)
                right_cell = RLImage(bildpfad, width=1.5 * inch, height=1.1 * inch)

        header_tbl = Table([[title_para, right_cell]], colWidths=[5.0 * inch, 1.5 * inch], hAlign="LEFT")
        header_tbl.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ALIGN',  (1, 0), (1, 0), 'RIGHT'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        story.append(header_tbl)
        story.append(Spacer(1, 12))

        # Page 1: Filters & values
        story.append(Paragraph("Selected Filters:", styles["Heading2"]))
        filter_data = [["Criterion", "Value"]] + [[k, v] for k, v in filters.items()]
        table = Table(filter_data, hAlign='LEFT')
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
        ]))
        story.append(table)
        story.append(Spacer(1, 12))

        story.append(Paragraph("Driver Analysis Values:", styles["Heading2"]))
        if values_table_df is not None and not values_table_df.empty:
            headers = list(values_table_df.columns)
            values_table_data = [headers] + values_table_df.values.tolist()
        else:
            values_table_data = [["Criterion", "Value"]] + chart_df.values.tolist()

        val_table = Table(values_table_data, hAlign='LEFT')
        val_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        story.append(val_table)
        story.append(PageBreak())

        # Page 2: Chart
        story.append(Paragraph("Driver Analysis (Chart):", styles["Heading2"]))
        if chart_png_bytes:
            img_buf = io.BytesIO(chart_png_bytes)
            story.append(RLImage(img_buf, width=6 * inch, height=4 * inch))
        else:
            story.append(Paragraph("Note: The chart could not be exported from the view.", styles["Italic"]))
        story.append(PageBreak())

        # Page 3: Model Fit & Dropped Drivers
        if adjusted_r2 is not None and n_value is not None:
            story.append(Paragraph("Model Fit:", styles["Heading2"]))
            story.append(Paragraph(f"Adjusted R²: {adjusted_r2}", styles["Normal"]))
            story.append(Paragraph(f"n: {n_value}", styles["Normal"]))
            story.append(Spacer(1, 12))

            story.append(Paragraph("Dropped Significant Drivers:", styles["Heading2"]))
            if dropped_drivers:
                dd_tbl = Table([["Criterion"]] + [[d] for d in dropped_drivers], hAlign='LEFT')
                dd_tbl.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
                ]))
                story.append(dd_tbl)
            else:
                story.append(Paragraph("No 'Dropped Drivers' available.", styles["Normal"]))

        doc.build(story)
        buffer.seek(0)
        return buffer

    def create_pdf_and_download_link(chart_png_bytes):
        # Filter values
        filter_values = {
            "Project": projekt,
            "Data Group": datengruppe,
            "Country": land,
            "Psychography": psychographie,
            "Stage": stage,
            "Market": markt,
            "Driver Set": treiberset
        }
        if datengruppe == "Cluster":
            filter_values["Cluster_Info"] = clusterinfo if clusterinfo is not None else "All"

        # Dynamic value-table columns (wie gehabt)
        dynamic_cols = []
        if land == "All": dynamic_cols.append("Country")
        if psychographie == "All": dynamic_cols.append("Psychography")
        if stage == "All": dynamic_cols.append("Stage")
        if markt == "All": dynamic_cols.append("Market")
        if treiberset == "All": dynamic_cols.append("Driver Set")
        if datengruppe == "Cluster" and (clusterinfo is None or clusterinfo == "All"):
            if "Cluster_Info" in df.columns:
                dynamic_cols.append("Cluster_Info")

        base_cols = ["Entity", "Value"]
        table_cols = base_cols + dynamic_cols
        values_table_df = df[table_cols].copy() if table_cols else df[base_cols].copy()

        chart_data = df[["Entity", "Value"]]

        alle_filter_gesetzt_local = all([
           land != "All",
           psychographie != "All",
           (stage != "All" if general_factor == "No" else True),
           markt != "All",
           treiberset != "All",
           (locals().get("clusterinfo", "All") != "All" if datengruppe == "Cluster" else True)
        ])

        adj_r2 = df["Adjusted_R2"].iloc[0] if ("Adjusted_R2" in df.columns and not df.empty and alle_filter_gesetzt_local) else None
        n_val  = df["n"].iloc[0]           if ("n"           in df.columns and not df.empty and alle_filter_gesetzt_local) else None
        dropped = df["Dropped_Drivers"].iloc[0] if ("Dropped_Drivers" in df.columns and not df.empty and alle_filter_gesetzt_local) else None

        pdf_buffer = export_to_pdf(
            filters=filter_values,
            chart_df=chart_data,
            chart_png_bytes=chart_png_bytes,
            adjusted_r2=adj_r2,
            n_value=n_val,
            dropped_drivers=dropped,
            values_table_df=values_table_df
        )

        b64 = base64.b64encode(pdf_buffer.read()).decode()
        href = f'''
            <a href="data:application/pdf;base64,{b64}" download="dashboard_export.pdf" style="text-decoration: none;">
                <button style="
                    padding: 6px 24px;
                    font-size: 14px;
                    background-color: #444;
                    color: white;
                    border: 1px solid #222;
                    border-radius: 6px;
                    cursor: pointer;
                    margin-top: 2px;
                ">📄 Export this view to PDF</button>
            </a>
        '''
        return href

    # ==========================================================
    # === Reihenfolge gemäß deinem Wunsch (unter dem Titel)  ===
    # ==========================================================

    # 1) Chart-Objekt erstellen & PNG ERZEUGEN (noch nicht anzeigen!)
    chart_obj = None
    chart_png_bytes = None
    if not df.empty:
        chart_obj = alt.Chart(df).mark_bar(
            color=farbe,
            stroke='#333',
            strokeWidth=1.5
        ).encode(
            x=alt.X("Value:Q", title="Value"),
            y=alt.Y("Entity:N", sort='-x', title="Criterion"),
            tooltip=["Entity", "Value"]
        ).properties(width=700, height=400)

        # PNG genau aus diesem Chart erzeugen -> garantiert "aktuelle Ansicht"
        try:
            buf = io.BytesIO()
            save(chart_obj, buf, fmt="png")
            buf.seek(0)
            chart_png_bytes = buf.getvalue()
        except Exception:
            try:
                chart_png_bytes = vlc.vegalite_to_png(chart_obj.to_dict())
            except Exception:
                chart_png_bytes = None

    # 2) Export-Link bauen (nutzt DIREKT chart_png_bytes dieses Laufs)
    export_link_html = create_pdf_and_download_link(chart_png_bytes)

    # 3) Header-Zeile (Results View | Psychographie-Bild | Export-Button) DIREKT unter dem Titel
    header_col1, header_col2, header_col3 = st.columns([2, 4, 1])
    with header_col1:
        st.write("## Results View")
    with header_col2:
        if psychographie != "All":
            bildpfad = f"graphics/{psychographie}.png"
            if os.path.exists(bildpfad):
                header_col2.image(bildpfad, width=100)
    with header_col3:
        st.markdown(export_link_html, unsafe_allow_html=True)

    # 4) Customer Journey direkt UNTER der Header-Zeile und ÜBER dem Diagramm
    def render_chevrons(selected_stage, farbe, psychographie):
        html = "<div style='display:flex; gap:12px; margin-bottom:24px;'>"
        for step in JOURNEY_STEPS:
            if psychographie != "All" and selected_stage is None:
                color = farbe
            elif selected_stage:
                if psychographie != "All":
                    color = farbe if step == selected_stage else "#888"
                else:
                    color = "#333" if step == selected_stage else "#888"
            else:
                color = "#888"
            text_color = "white"
            html += f"<div style='background:{color}; padding:14px 60px; border-radius:4px; font-weight:bold; color:{text_color};'>{step}</div>"
        html += "</div>"
        return html

    st.markdown(render_chevrons(selected_stage, farbe, psychographie), unsafe_allow_html=True)

    # 5) Diagramm JETZT anzeigen (steht dadurch unter Header + Journey)
    if chart_obj is not None:
        st.altair_chart(chart_obj, use_container_width=True)

    # === Model fit and dropped drivers display (unverändert) ===
    alle_filter_gesetzt = all([
        land != "All",
        psychographie != "All",
        stage != "All",
        markt != "All",
        treiberset != "All",
        (clusterinfo != "All" if datengruppe == "Cluster" else True)
    ])

    if alle_filter_gesetzt and not df.empty:
        st.write("### Model Fit")
        col1, col2 = st.columns(2)
        col1.metric("Adjusted R²", df["Adjusted_R2"].iloc[0])
        col2.metric("n", df["n"].iloc[0])

        st.write("### Dropped Significant Drivers")
        dropped = df["Dropped_Drivers"].iloc[0] if isinstance(df["Dropped_Drivers"].iloc[0], list) else []
        if dropped:
            st.table(pd.DataFrame(dropped, columns=["Dropped Driver"]))
        else:
            st.info("No 'Dropped Drivers' found in this file.")

else:
    st.write("### Please select a project and a data group first.")

