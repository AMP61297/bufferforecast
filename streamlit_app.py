import streamlit as st
import pandas as pd
import numpy as np
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import io
import xlsxwriter
import datetime
import zoneinfo


# ✅ CSS einfügen, um horizontales Scrollen in Tabellen zu vermeiden
st.markdown("""
    <style>
    .stDataFrame div[data-testid="stHorizontalBlock"] {
        overflow-x: auto;
        white-space: nowrap;
    }
    .stDataFrame th, .stDataFrame td {
        min-width: 100px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)
# --- Kopfzeile & Logo ---
st.title("📦 Pufferprognose")
st.markdown("#### Bereich: Vormontage")


# --- Sidebar für Einstellungen ---
st.sidebar.header("Grenzwerte für Puffer Ende")
min_grenze = st.sidebar.slider("🔽 Mindestwert", min_value=0, max_value=50, value=8)
max_grenze = st.sidebar.slider("🔼 Maximalwert", min_value=10, max_value=100, value=60)

st.sidebar.header("Zeitraum einstellen")
start_datum = st.sidebar.date_input("Startdatum", value=datetime.date.today())
anzeige_tage = st.sidebar.slider("Anzahl Tage anzeigen", min_value=1, max_value=30, value=5)
tage = [start_datum + datetime.timedelta(days=i) for i in range(anzeige_tage)]

# --- Linien definieren ---
linien = ["Linie 2", "Linie 3", "Linie 14", "Linie 15"]
ausgewaehlte_linien = st.sidebar.multiselect("Wähle Montagelinien", linien, default=linien)

# --- Dummy-Daten erstellen ---
data = []
for linie in linien:
    for tag in tage:
        data.append({
            "Linie": linie,
            "Datum": tag,
            "Puffer Start": np.nan,
            "Zulauf": np.nan,
            "Ablauf": np.nan,
            "Ausschleuser": np.nan
        })

df_input = pd.DataFrame(data)
df_input = df_input[df_input["Linie"].isin(ausgewaehlte_linien)].copy()
df_edited = df_input.copy()


# --- Berechnung Puffer Ende (mit 93% Zulauf) ---
ZULAUF_FAKTOR = 0.93
df_edited["Zulauf berechnet (93 %)"] = np.nan
df_edited["Puffer Ende"] = np.nan
df_edited = df_edited.sort_values(["Linie", "Datum"]).reset_index(drop=True)

for linie in df_edited["Linie"].unique():
    df_linie = df_edited[df_edited["Linie"] == linie].copy()

    for i in range(len(df_linie)):
        zulauf = df_linie.iloc[i]["Zulauf"]
        ablauf = df_linie.iloc[i]["Ablauf"]
        effektiver_zulauf = round(zulauf * ZULAUF_FAKTOR) if pd.notna(zulauf) else np.nan
        df_linie.loc[df_linie.index[i], "Zulauf berechnet (93 %)"] = effektiver_zulauf

        if i == 0:
            if pd.notna(df_linie.iloc[i]["Puffer Start"]) and pd.notna(effektiver_zulauf) and pd.notna(ablauf):
                df_linie.loc[df_linie.index[i], "Puffer Ende"] = (
                    df_linie.iloc[i]["Puffer Start"] + effektiver_zulauf - ablauf
                )
        else:
            if pd.isna(df_linie.iloc[i]["Puffer Start"]):
                df_linie.loc[df_linie.index[i], "Puffer Start"] = df_linie.iloc[i - 1]["Puffer Ende"]
            if pd.notna(df_linie.iloc[i]["Puffer Start"]) and pd.notna(effektiver_zulauf) and pd.notna(ablauf):
                df_linie.loc[df_linie.index[i], "Puffer Ende"] = (
                    df_linie.iloc[i]["Puffer Start"] + effektiver_zulauf - ablauf
                )

    df_edited.update(df_linie)

# --- Spaltenreihenfolge ---
df_edited = df_edited[[
    "Linie", "Datum", "Puffer Start", "Zulauf", "Zulauf berechnet (93 %)",
    "Ablauf", "Ausschleuser", "Puffer Ende"
]]

# ✅ Leere Zellen durch 0 ersetzen
df_edited.fillna(0, inplace=True)

# Zwei nebeneinanderstehende Tabellen – editierbar & berechnet
col1, col2 = st.columns(2)

with col1:
    st.subheader("✏️ Eingabedaten")
    st.data_editor(
        df_input.fillna(0),
        use_container_width=True,
        height=anzeige_tage * 43 + 100,  # dynamisch auf Zeilenhöhe abgestimmt
        num_rows="dynamic",
        disabled=["Linie", "Datum"]
    )

with col2:
    st.subheader("📋 Berechnete Puffer Ende")
    st.dataframe(
        df_edited.fillna(0),
        use_container_width=True,
        height=anzeige_tage * 43 + 100  # gleiche Höhe wie links
    )


# --- Diagramm über die volle Breite anzeigen ---
st.subheader("📊 Verlauf Puffer Ende")
fig, ax = plt.subplots(figsize=(18, 6))  # ⬅️ Größer gemacht

for linie in ausgewaehlte_linien:
    df_plot = df_edited[df_edited["Linie"] == linie]
    ax.plot(df_plot["Datum"], df_plot["Puffer Ende"], marker="o", label=linie, linewidth=2)

ax.axhspan(min_grenze, max_grenze, color='lightgreen', alpha=0.4, label="Zielbereich")
ax.set_facecolor('#FAFAFA')
fig.patch.set_facecolor('#FAFAFA')
ax.set_title("📊 Entwicklung Puffer Ende", fontsize=16)
ax.set_xlabel("Datum")
ax.set_ylabel("Pufferbestand")
ax.xaxis.set_major_formatter(mdates.DateFormatter('%d.%m'))
ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
fig.autofmt_xdate(rotation=45)
ax.grid(True, linestyle="--", alpha=0.3)
ax.legend()

st.pyplot(fig)

# --- Excel-Export mit Bild, Zeitstempel & Formatierung ---
berlin = zoneinfo.ZoneInfo("Europe/Berlin")
zeitstempel = datetime.datetime.now(berlin).strftime("Exportzeitpunkt: %Y-%m-%d %H:%M:%S")

output = io.BytesIO()
image_path = "puffer_chart.png"
fig.savefig(image_path, bbox_inches='tight')

with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_edited.to_excel(writer, index=False, sheet_name="Pufferprognose", startrow=2)

    workbook = writer.book
    worksheet = writer.sheets["Pufferprognose"]

    # 📅 Zeitstempel in Zelle L1
    worksheet.write("L1", zeitstempel)

    # 📊 Bild unter der Tabelle einfügen
    image_row = len(df_edited) + 5
    worksheet.insert_image(image_row, 0, image_path, {
        'x_offset': 0, 'y_offset': 10,
        'x_scale': 1.0, 'y_scale': 1.0
    })

    # 📐 Formatierte Tabelle
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'middle',
        'fg_color': '#F36F21',
        'color': 'white',
        'border': 1
    })

    for col_num, value in enumerate(df_edited.columns.values):
        worksheet.write(2, col_num, value, header_format)

    table_range = f"A3:H{len(df_edited) + 3}"
    worksheet.add_table(table_range, {
        'columns': [{'header': col} for col in df_edited.columns],
        'style': 'Table Style Medium 9'
    })

    # 📏 Spaltenbreite automatisch anpassen
    for i, column in enumerate(df_edited.columns):
        max_len = max(df_edited[column].astype(str).map(len).max(), len(column)) + 2
        worksheet.set_column(i, i, max_len)

    # 📄 Seitenlayout
    worksheet.set_paper(9)
    worksheet.fit_to_pages(1, 0)
    worksheet.center_horizontally()
    worksheet.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)

# 📥 Download-Button
st.download_button(
    label="📥 Excel-Datei herunterladen",
    data=output.getvalue(),
    file_name="pufferprognose.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


