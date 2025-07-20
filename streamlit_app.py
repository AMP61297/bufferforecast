import streamlit as st
import pandas as pd
import numpy as np
import datetime
import matplotlib.pyplot as plt
import matplotlib.dates as mdates



st.set_page_config(page_title="Bufferforecast", layout="wide")
st.title("📦 Pufferprognose HGM")

# --- Sidebar für Min/Max ---
st.sidebar.header("Grenzwerte für Puffer Ende")
min_grenze = st.sidebar.slider("🔽 Mindestwert", min_value=0, max_value=50, value=8)
max_grenze = st.sidebar.slider("🔼 Maximalwert", min_value=10, max_value=100, value=60)

# --- Zeitraum wählbar machen ---
st.sidebar.header("Zeitraum einstellen")
start_datum = st.sidebar.date_input("Startdatum", value=datetime.date.today())
anzeige_tage = st.sidebar.slider("Anzahl Tage anzeigen", min_value=1, max_value=30, value=10)
tage = [start_datum + datetime.timedelta(days=i) for i in range(anzeige_tage)]

# --- Linien definieren ---
linien = ["Linie 2", "Linie 3", "Linie 14", "Linie 15"]

# --- Dummy-Daten ---
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

# --- Linienauswahl ---
ausgewaehlte_linien = st.multiselect("Wähle Montagelinien", linien, default=linien)
df_input = df_input[df_input["Linie"].isin(ausgewaehlte_linien)].copy()

# --- Eingabemaske ---
df_edited = st.data_editor(
    df_input,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "Linie": st.column_config.TextColumn(disabled=True),
        "Datum": st.column_config.DateColumn(disabled=True),
        "Puffer Start": st.column_config.NumberColumn(),
        "Zulauf": st.column_config.NumberColumn(),
        "Ablauf": st.column_config.NumberColumn(),
        "Ausschleuser": st.column_config.NumberColumn()
    }
)

# Sicherstellen, dass "Puffer Ende" existiert
df_edited["Puffer Ende"] = np.nan

# --- Puffer Ende berechnen mit 93% effektivem Zulauf (gerundet) + Anzeige der berechneten Spalte ---
df_edited = df_edited.sort_values(["Linie", "Datum"]).reset_index(drop=True)

ZULAUF_FAKTOR = 0.93
df_edited["Zulauf berechnet (93 %)"] = np.nan

for linie in df_edited["Linie"].unique():
    df_linie = df_edited[df_edited["Linie"] == linie].copy()

    for i in range(len(df_linie)):
        zulauf = df_linie.iloc[i]["Zulauf"]
        ablauf = df_linie.iloc[i]["Ablauf"]

        # Berechne korrigierten Zulauf (gerundet)
        effektiver_zulauf = round(zulauf * ZULAUF_FAKTOR) if pd.notna(zulauf) else np.nan
        df_linie.loc[df_linie.index[i], "Zulauf berechnet (93 %)"] = effektiver_zulauf

        if i == 0:
            if (
                pd.notna(df_linie.iloc[i]["Puffer Start"])
                and pd.notna(effektiver_zulauf)
                and pd.notna(ablauf)
            ):
                df_linie.loc[df_linie.index[i], "Puffer Ende"] = (
                    df_linie.iloc[i]["Puffer Start"] + effektiver_zulauf - ablauf
                )
        else:
            if pd.isna(df_linie.iloc[i]["Puffer Start"]):
                df_linie.loc[df_linie.index[i], "Puffer Start"] = df_linie.iloc[i - 1]["Puffer Ende"]
            if (
                pd.notna(df_linie.iloc[i]["Puffer Start"])
                and pd.notna(effektiver_zulauf)
                and pd.notna(ablauf)
            ):
                df_linie.loc[df_linie.index[i], "Puffer Ende"] = (
                    df_linie.iloc[i]["Puffer Start"] + effektiver_zulauf - ablauf
                )

    df_edited.update(df_linie)

# --- Nach Zeitraum filtern ---
df_edited = df_edited[df_edited["Datum"].between(start_datum, start_datum + datetime.timedelta(days=anzeige_tage - 1))]

# 🔄 Spalten sortieren für bessere Anzeige
spalten_sortiert = [
    "Linie",
    "Datum",
    "Puffer Start",
    "Zulauf",
    "Zulauf berechnet (93 %)",
    "Ablauf",
    "Ausschleuser",
    "Puffer Ende"
]
df_edited = df_edited[spalten_sortiert]


# --- Nach Zeitraum filtern ---
df_edited = df_edited[df_edited["Datum"].between(start_datum, start_datum + datetime.timedelta(days=anzeige_tage - 1))]

# --- Tabelle anzeigen ---
st.subheader("📋 Tabelle mit berechnetem Puffer Ende")
st.dataframe(df_edited, use_container_width=True)

# --- Diagramm anzeigen ---
st.subheader("📊 Verlauf Puffer Ende")

fig, ax = plt.subplots(figsize=(10, 5))

for linie in ausgewaehlte_linien:
    df_plot = df_edited[df_edited["Linie"] == linie]
    ax.plot(df_plot["Datum"], df_plot["Puffer Ende"], marker="o", label=linie)

# --- Zielbereich visualisieren ---
ax.axhspan(min_grenze, max_grenze, color='lightgreen', alpha=0.3, label="Zielbereich")

ax.set_xlabel("Datum")
ax.set_ylabel("Puffer Ende")
ax.set_title("Pufferentwicklung je Linie")
ax.xaxis.set_major_formatter(mdates.DateFormatter('%d.%m'))
ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
fig.autofmt_xdate(rotation=45)
ax.grid(True)
ax.legend()

st.pyplot(fig)

import io
import xlsxwriter
from datetime import datetime
import zoneinfo  # ⬅️ Zeitzonenimport nach datetime

# 🔽 Direkt danach kommt dein Zeitstempel in Berliner Zeit
berlin = zoneinfo.ZoneInfo("Europe/Berlin")
zeitstempel = datetime.now(berlin).strftime("Exportzeitpunkt: %Y-%m-%d %H:%M:%S")


# 🔄 Excel-Datei vorbereiten
output = io.BytesIO()
zeitstempel = datetime.now().strftime("Exportzeitpunkt: %Y-%m-%d %H:%M:%S")

with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
    df_edited.to_excel(writer, index=False, sheet_name="Pufferprognose", startrow=2)  # Tabelle beginnt in Zeile 3

    workbook = writer.book
    worksheet = writer.sheets["Pufferprognose"]

    # 🕒 Zeitstempel in Zelle L1 (rechte Seite)
    worksheet.write("L1", zeitstempel)

    # 🖼️ Bild speichern
    image_path = "puffer_chart.png"
    fig.savefig(image_path, bbox_inches='tight')

    # 📊 Bild unterhalb der Tabelle einfügen
    image_row = len(df_edited) + 5
    worksheet.insert_image(image_row, 0, image_path, {
        'x_offset': 0,
        'y_offset': 10,
        'x_scale': 1.0,
        'y_scale': 1.0
    })

    # 📄 Seitenlayout optimieren
    worksheet.set_paper(9)  # A4
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


