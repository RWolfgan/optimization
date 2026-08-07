import streamlit as st
import pandas as pd
import pulp
import io

# Seitenkonfiguration
st.set_page_config(page_title="4D Top 10 Optimierer", page_icon="📊", layout="wide")

st.title("🧮 4D Zielwert-Optimierer (Top 10 Varianten)")
st.write("Laden Sie Ihre Excel-Liste hoch. Das System berechnet die 10 besten Kombinationen (sortiert nach der geringsten Gesamtabweichung).")

# ==========================================
# 1. TEMPLATE-DOWNLOAD FÜR DIE STRUKTUR
# ==========================================
st.sidebar.header("1. Struktur-Vorlage")

template_df = pd.DataFrame({
    "Objektname": [f"Objekt_{i}" for i in range(1, 41)],
    "Jahr_1": [50] * 40,
    "Jahr_2": [50] * 40,
    "Jahr_3": [50] * 40,
    "Jahr_4": [50] * 40
})

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
        df = pd.read_excel(uploaded_file)
        erforderliche_spalten = ["Objektname", "Jahr_1", "Jahr_2", "Jahr_3", "Jahr_4"]
        
        if not all(spalte in df.columns for spalte in erforderliche_spalten):
            st.error(f"❌ Fehler: Die Excel-Datei muss exakt die Spalten enthalten: {', '.join(erforderliche_spalten)}")
        else:
            st.success("✅ Excel-Struktur erfolgreich erkannt!")
            
            if st.button("🚀 Top 10 Kombinationen berechnen"):
                objekte = df["Objektname"].tolist()
                j1_werte = dict(zip(df["Objektname"], df["Jahr_1"]))
                j2_werte = dict(zip(df["Objektname"], df["Jahr_2"]))
                j3_werte = dict(zip(df["Objektname"], df["Jahr_3"]))
                j4_werte = dict(zip(df["Objektname"], df["Jahr_4"]))
                
                # Listen zum Speichern der Ergebnisse
                gefundene_varianten = []  # Speichert die 0/1 Verteilung pro Durchlauf
                varianten_metriken = []   # Speichert die Summen und Abweichungen
                
                # Wir suchen nach maximal 10 unterschiedlichen Lösungen
                anzahl_varianten = 10
                
                with st.spinner("Berechne mathematische Varianten..."):
                    for v_idx in range(1, anzahl_varianten + 1):
                        
                        # Neues Optimierungsproblem aufsetzen
                        prob = pulp.LpProblem(f"Abweichung_Variante_{v_idx}", pulp.LpMinimize)
                        
                        # Entscheidungsvariablen
                        waehle = pulp.LpVariable.dicts("Waehle", objekte, cat='Binary')
                        
                        # Hilfsvariablen für Abweichungen
                        abw_pos_j1 = pulp.LpVariable("Abw_Pos_J1", lowBound=0)
                        abw_neg_j1 = pulp.LpVariable("Abw_Neg_J1", lowBound=0)
                        abw_pos_j2 = pulp.LpVariable("Abw_Pos_J2", lowBound=0)
                        abw_neg_j2 = pulp.LpVariable("Abw_Neg_J2", lowBound=0)
                        abw_pos_j3 = pulp.LpVariable("Abw_Pos_J3", lowBound=0)
                        abw_neg_j3 = pulp.LpVariable("Abw_Neg_J3", lowBound=0)
                        abw_pos_j4 = pulp.LpVariable("Abw_Pos_J4", lowBound=0)
                        abw_neg_j4 = pulp.LpVariable("Abw_Neg_J4", lowBound=0)
                        
                        # Zielfunktion: Minimiere Gesamtabweichung
                        prob += (abw_pos_j1 + abw_neg_j1 + abw_pos_j2 + abw_neg_j2 + 
                                 abw_pos_j3 + abw_neg_j3 + abw_pos_j4 + abw_neg_j4)
                        
                        # Jahres-Bedingungen
                        prob += pulp.lpSum([j1_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j1 + abw_neg_j1 == ziel_j1
                        prob += pulp.lpSum([j2_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j2 + abw_neg_j2 == ziel_j2
                        prob += pulp.lpSum([j3_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j3 + abw_neg_j3 == ziel_j3
                        prob += pulp.lpSum([j4_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j4 + abw_neg_j4 == ziel_j4
                        
                        # ZUSTÄTZLICHE BEDINGUNG: Bereits gefundene Kombinationen ausschließen
                        for vorherige_sol in gefundene_varianten:
                            # Erzwingt, dass sich mindestens eine Entscheidung (0 zu 1 oder 1 zu 0) ändern muss
                            prob += pulp.lpSum([waehle[obj] if vorherige_sol[obj] == 1 else (1 - waehle[obj]) for obj in objekte]) <= len(objekte) - 1
                        
                        # Solver starten
                        prob.solve(pulp.PULP_CBC_CMD(msg=False))
                        
                        # Wenn eine gültige Lösung gefunden wurde
                        if pulp.LpStatus[prob.status] == "Optimal":
                            # Aktuelle Lösung sichern
                            aktuelle_loesung = {obj: int(waehle[obj].varValue) for obj in objekte}
                            gefundene_varianten.append(aktuelle_loesung)
                            
                            # Werte für die Übersicht berechnen
                            sum_j1 = sum(j1_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                            sum_j2 = sum(j2_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                            sum_j3 = sum(j3_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                            sum_j4 = sum(j4_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                            total_abw = pulp.value(prob.objective)
                            
                            varianten_metriken.append({
                                "Variante": f"Variante_{v_idx}",
                                "Abweichung Gesamt": total_abw,
                                "Jahr 1 Summe": sum_j1,
                                "Jahr 2 Summe": sum_j2,
                                "Jahr 3 Summe": sum_j3,
                                "Jahr 4 Summe": sum_j4
                            })
                        else:
                            # Keine weiteren mathematischen Kombinationen mehr möglich
                            break
                
                # ==========================================
                # 4. AUSGABE DER ERGEBNISSE IM INTERFACE
                # ==========================================
                if len(gefundene_varianten) == 0:
                    st.error("Es konnte keine Kombination berechnet werden.")
                else:
                    st.success(f"Erfolgreich {len(gefundene_varianten)} Varianten berechnet!")
                    
                    # Metriken-Tabelle anzeigen
                    st.write("### 📈 Übersicht der berechneten Varianten (Sortiert nach bester Annäherung):")
                    metriken_df = pd.DataFrame(varianten_metriken)
                    st.dataframe(metriken_df, use_container_width=True)
                    
                    # Die neuen Spalten in den Haupt-DataFrame eintragen
                    for idx, sol in enumerate(gefundene_varianten):
                        v_name = f"Variante_{idx+1} (1=Ja)"
                        df[v_name] = [sol[obj] for obj in df["Objektname"]]
                    
                    st.write("### 📋 Daten-Vorschau inklusive aller 10 Varianten-Spalten:")
                    st.dataframe(df)
                    
                    # ==========================================
                    # 5. EXCEL EXPORT MIT DEN 10 SPALTEN
                    # ==========================================
                    buffer_export = io.BytesIO()
                    with pd.ExcelWriter(buffer_export, engine='openpyxl') as writer:
                        # Tabellenblatt 1: Die Hauptdaten mit den 10 Entscheidungsspalten
                        df.to_excel(writer, index=False, sheet_name="Objekt_Auswahl")
                        # Tabellenblatt 2: Die Zusammenfassung der Abweichungen
                        metriken_df.to_excel(writer, index=False, sheet_name="Varianten_Vergleich")
                        
                    st.download_button(
                        label="📥 Excel mit allen 10 Varianten herunterladen",
                        data=buffer_export.getvalue(),
                        file_name="4d_optimierung_top10.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung: {e}")
