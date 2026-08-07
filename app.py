import streamlit as st
import pandas as pd
import pulp
import io

# Seitenkonfiguration
st.set_page_config(page_title="4D Binär-Optimierer", page_icon="📊", layout="wide")

st.title("🧮 4D Multi-Periodischer Zielwert-Optimierer")
st.write("Laden Sie eine Excel-Liste mit 40 Objekten hoch, um die beste Kombination für Ihre 4 Jahres-Zielsummen zu finden.")

# ==========================================
# 1. TEMPLATE-DOWNLOAD FÜR DIE STRUKTUR
# ==========================================
st.sidebar.header("1. Struktur-Vorlage")
st.sidebar.write("Die Excel-Datei muss genau diese 5 Spalten enthalten:")

# Beispiel-Datenrahmen für die Vorlage erstellen
template_df = pd.DataFrame({
    "Objektname": [f"Objekt_{i}" for i in range(1, 41)],
    "Jahr_1": [50] * 40,
    "Jahr_2": [60] * 40,
    "Jahr_3": [70] * 40,
    "Jahr_4": [80] * 40
})

# In Excel-Buffer schreiben
buffer_template = io.BytesIO()
with pd.ExcelWriter(buffer_template, engine='openpyxl') as writer:
    template_df.to_excel(writer, index=False, sheet_name="Daten")

st.sidebar.download_button(
    label="📄 Excel-Vorlage herunterladen",
    data=buffer_template.getvalue(),
    file_name="struktur_vorlage_4d.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ==========================================
# 2. ZIELWERTE ABFRAGEN
# ==========================================
st.sidebar.header("2. Gewünschte Zielsummen")
ziel_j1 = st.sidebar.number_input("Ziel Gesamtsumme Jahr 1", value=1200, step=1)
ziel_j2 = st.sidebar.number_input("Ziel Gesamtsumme Jahr 2", value=1100, step=1)
ziel_j3 = st.sidebar.number_input("Ziel Gesamtsumme Jahr 3", value=1300, step=1)
ziel_j4 = st.sidebar.number_input("Ziel Gesamtsumme Jahr 4", value=1000, step=1)

# ==========================================
# 3. DATEI-UPLOAD & VERARBEITUNG
# ==========================================
uploaded_file = st.file_uploader("Wählen Sie Ihre ausgefüllte Excel-Datei aus", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Excel einlesen
        df = pd.read_excel(uploaded_file)
        
        # Validierung der Spalten
        erforderliche_spalten = ["Objektname", "Jahr_1", "Jahr_2", "Jahr_3", "Jahr_4"]
        if not all(spalte in df.columns for spalte in erforderliche_spalten):
            st.error(f"❌ Fehler: Die Excel-Datei muss exakt die Spalten enthalten: {', '.join(erforderliche_spalten)}")
        else:
            st.success("✅ Excel-Struktur erfolgreich erkannt!")
            st.write("### Vorschau Ihrer Daten (Erste Zeilen):")
            st.dataframe(df.head())
            
            # Button zum Starten der Optimierung
            if st.button("🚀 Optimale Kombination berechnen"):
                
                # Daten für PuLP vorbereiten
                objekte = df["Objektname"].tolist()
                j1_werte = dict(zip(df["Objektname"], df["Jahr_1"]))
                j2_werte = dict(zip(df["Objektname"], df["Jahr_2"]))
                j3_werte = dict(zip(df["Objektname"], df["Jahr_3"]))
                j4_werte = dict(zip(df["Objektname"], df["Jahr_4"]))
                
                # Optimierungsproblem aufsetzen (Minimierung der Abweichung)
                prob = pulp.LpProblem("Minimale_Abweichung", pulp.LpMinimize)
                
                # Entscheidungsvariablen (Binär: 0 oder 1)
                waehle = pulp.LpVariable.dicts("Waehle", objekte, cat='Binary')
                
                # Hilfsvariablen für absolute Abweichungen (Positiv / Negativ)
                abw_pos_j1 = pulp.LpVariable("Abw_Pos_J1", lowBound=0)
                abw_neg_j1 = pulp.LpVariable("Abw_Neg_J1", lowBound=0)
                abw_pos_j2 = pulp.LpVariable("Abw_Pos_J2", lowBound=0)
                abw_neg_j2 = pulp.LpVariable("Abw_Neg_J2", lowBound=0)
                abw_pos_j3 = pulp.LpVariable("Abw_Pos_J3", lowBound=0)
                abw_neg_j3 = pulp.LpVariable("Abw_Neg_J3", lowBound=0)
                abw_pos_j4 = pulp.LpVariable("Abw_Pos_J4", lowBound=0)
                abw_neg_j4 = pulp.LpVariable("Abw_Neg_J4", lowBound=0)
                
                # Zielfunktion: Minimiere die Gesamtabweichung
                prob += (abw_pos_j1 + abw_neg_j1 + abw_pos_j2 + abw_neg_j2 + 
                         abw_pos_j3 + abw_neg_j3 + abw_pos_j4 + abw_neg_j4)
                
                # Bedingungen für die Fehlerberechnung pro Jahr
                prob += pulp.lpSum([j1_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j1 + abw_neg_j1 == ziel_j1
                prob += pulp.lpSum([j2_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j2 + abw_neg_j2 == ziel_j2
                prob += pulp.lpSum([j3_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j3 + abw_neg_j3 == ziel_j3
                prob += pulp.lpSum([j4_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j4 + abw_neg_j4 == ziel_j4
                
                # Solver ausführen
                prob.solve(pulp.PULP_CBC_CMD(msg=False))
                
                # Ergebnisse auswerten
                if pulp.LpStatus[prob.status] == "Optimal":
                    # Entscheidungs-Spalte dem DataFrame hinzufügen
                    df["Ausgewählt (1=Ja, 0=Nein)"] = [int(waehle[obj].varValue) for obj in df["Objektname"]]
                    
                    # Berechnete Jahressummen ermitteln
                    sum_j1 = sum(df[df["Ausgewählt (1=Ja, 0=Nein)"] == 1]["Jahr_1"])
                    sum_j2 = sum(df[df["Ausgewählt (1=Ja, 0=Nein)"] == 1]["Jahr_2"])
                    sum_j3 = sum(df[df["Ausgewählt (1=Ja, 0=Nein)"] == 1]["Jahr_3"])
                    sum_j4 = sum(df[df["Ausgewählt (1=Ja, 0=Nein)"] == 1]["Jahr_4"])
                    
                    # Feedback im UI anzeigen
                    st.write("### 📊 Optimierungsergebnis")
                    col1, col2, col3, col4 = st.columns(4)
                    col1.metric("Jahr 1 Summe", f"{sum_j1}", f"Ziel: {ziel_j1}")
                    col2.metric("Jahr 2 Summe", f"{sum_j2}", f"Ziel: {ziel_j2}")
                    col3.metric("Jahr 3 Summe", f"{sum_j3}", f"Ziel: {ziel_j3}")
                    col4.metric("Jahr 4 Summe", f"{sum_j4}", f"Ziel: {ziel_j4}")
                    
                    # Filterung für die Anzeige im Web
                    ausgewaehlt_df = df[df["Ausgewählt (1=Ja, 0=Nein)"] == 1]
                    st.write(f"Es wurden **{len(ausgewaehlt_df)}** Objekte ausgewählt.")
                    st.dataframe(df)
                    
                    # ==========================================
                    # EXPORT ALS EXCEL ERSTELLEN
                    # ==========================================
                    buffer_export = io.BytesIO()
                    with pd.ExcelWriter(buffer_export, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Optimierungsergebnis")
                        
                    st.download_button(
                        label="📥 Optimierte Excel-Liste herunterladen",
                        data=buffer_export.getvalue(),
                        file_name="optimiertes_ergebnis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Der Solver konnte keine valide Lösung berechnen.")
    except Exception as e:
        st.error(f"Fehler beim Lesen der Datei: {e}")
