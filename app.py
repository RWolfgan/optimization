import streamlit as st
import pandas as pd
import pulp
import io

# Seitenkonfiguration
st.set_page_config(page_title="4D Top 10 Optimierer", page_icon="📊", layout="wide")

st.title("🧮 Fast 4D Zielwert-Optimierer (Top 10 Varianten)")
st.write("Optimierte Version für größere Datenmengen (z.B. 60+ Zeilen) mit automatischer Zeitbegrenzung.")

# ==========================================
# 1. TEMPLATE-DOWNLOAD FÜR DIE STRUKTUR
# ==========================================
st.sidebar.header("1. Struktur-Vorlage")

template_df = pd.DataFrame({
    "Objektname": [f"Objekt_{i}" for i in range(1, 61)], # Vorlage jetzt direkt mit 60 Zeilen
    "Jahr_1": * 60,
    "Jahr_2": * 60,
    "Jahr_3": * 60,
    "Jahr_4": * 60
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
            st.success(f"✅ Excel-Struktur erfolgreich erkannt! ({len(df)} Zeilen gefunden)")
            
            if st.button("🚀 Top 10 Kombinationen blitzschnell berechnen"):
                objekte = df["Objektname"].tolist()
                j1_werte = dict(zip(df["Objektname"], df["Jahr_1"]))
                j2_werte = dict(zip(df["Objektname"], df["Jahr_2"]))
                j3_werte = dict(zip(df["Objektname"], df["Jahr_3"]))
                j4_werte = dict(zip(df["Objektname"], df["Jahr_4"]))
                
                gefundene_varianten = []
                varianten_metriken = []
                anzahl_varianten = 10
                
                # Fortschrittsanzeige für den Nutzer
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for v_idx in range(1, anzahl_varianten + 1):
                    status_text.text(f"Berechne Variante {v_idx} von 10... (Max. 5 Sek.)")
                    progress_bar.progress((v_idx - 1) / anzahl_varianten)
                    
                    prob = pulp.LpProblem(f"Abweichung_Variante_{v_idx}", pulp.LpMinimize)
                    waehle = pulp.LpVariable.dicts("Waehle", objekte, cat='Binary')
                    
                    abw_pos_j1 = pulp.LpVariable("Abw_Pos_J1", lowBound=0)
                    abw_neg_j1 = pulp.LpVariable("Abw_Neg_J1", lowBound=0)
                    abw_pos_j2 = pulp.LpVariable("Abw_Pos_J2", lowBound=0)
                    abw_neg_j2 = pulp.LpVariable("Abw_Neg_J2", lowBound=0)
                    abw_pos_j3 = pulp.LpVariable("Abw_Pos_J3", lowBound=0)
                    abw_neg_j3 = pulp.LpVariable("Abw_Neg_J3", lowBound=0)
                    abw_pos_j4 = pulp.LpVariable("Abw_Pos_J4", lowBound=0)
                    abw_neg_j4 = pulp.LpVariable("Abw_Neg_J4", lowBound=0)
                    
                    prob += (abw_pos_j1 + abw_neg_j1 + abw_pos_j2 + abw_neg_j2 + 
                             abw_pos_j3 + abw_neg_j3 + abw_pos_j4 + abw_neg_j4)
                    
                    prob += pulp.lpSum([j1_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j1 + abw_neg_j1 == ziel_j1
                    prob += pulp.lpSum([j2_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j2 + abw_neg_j2 == ziel_j2
                    prob += pulp.lpSum([j3_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j3 + abw_neg_j3 == ziel_j3
                    prob += pulp.lpSum([j4_werte[obj] * waehle[obj] for obj in objekte]) - abw_pos_j4 + abw_neg_j4 == ziel_j4
                    
                    for vorherige_sol in gefundene_varianten:
                        prob += pulp.lpSum([waehle[obj] if vorherige_sol[obj] == 1 else (1 - waehle[obj]) for obj in objekte]) <= len(objekte) - 1
                    
                    # NEU: timeLimit=5 setzt eine harte Grenze von 5 Sekunden pro Variante
                    # msg=False unterdrückt Konsolen-Logs
                    prob.solve(pulp.PULP_CBC_CMD(timeLimit=5, msg=False))
                    
                    # NEU: Wir akzeptieren "Optimal" ODER "Feasible" (ausreichend gut angenähert nach Zeitablauf)
                    solver_status = pulp.LpStatus[prob.status]
                    if solver_status in ["Optimal", "Feasible"] and waehle[objekte[0]].varValue is not None:
                        aktuelle_loesung = {obj: int(round(waehle[obj].varValue)) for obj in objekte}
                        gefundene_varianten.append(aktuelle_loesung)
                        
                        sum_j1 = sum(j1_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                        sum_j2 = sum(j2_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                        sum_j3 = sum(j3_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                        sum_j4 = sum(j4_werte[obj] * aktuelle_loesung[obj] for obj in objekte)
                        
                        # Absolute Abweichung manuell ausrechnen, falls Solver abgebrochen wurde
                        total_abw = abs(sum_j1 - ziel_j1) + abs(sum_j2 - ziel_j2) + abs(sum_j3 - ziel_j3) + abs(sum_j4 - ziel_j4)
                        
                        varianten_metriken.append({
                            "Variante": f"Variante_{v_idx}",
                            "Status": "Exakt" if solver_status == "Optimal" else "Gute Näherung (Zeitlimit)",
                            "Abweichung Gesamt": total_abw,
                            "Jahr 1 Summe": sum_j1,
                            "Jahr 2 Summe": sum_j2,
                            "Jahr 3 Summe": sum_j3,
                            "Jahr 4 Summe": sum_j4
                        })
                    else:
                        # Keine weiteren Lösungen mehr auffindbar
                        break
                
                # Fortschritt abschließen
                progress_bar.progress(1.0)
                status_text.text("Berechnung abgeschlossen!")
                
                # ==========================================
                # 4. AUSGABE DER ERGEBNISSE
                # ==========================================
                if len(gefundene_varianten) == 0:
                    st.error("Es konnte innerhalb des Zeitlimits leider keine sinnvolle Kombination berechnet werden. Bitte prüfen Sie Ihre Zielwerte.")
                else:
                    st.success(f"Erfolgreich {len(gefundene_varianten)} Varianten berechnet!")
                    
                    st.write("### 📈 Übersicht der berechneten Varianten:")
                    metriken_df = pd.DataFrame(varianten_metriken)
                    st.dataframe(metriken_df, use_container_width=True)
                    
                    # Dynamisch die Spalten für alle gefundenen Varianten anhängen
                    for idx, sol in enumerate(gefundene_varianten):
                        v_name = f"Variante_{idx+1} (1=Ja)"
                        df[v_name] = [sol[obj] for obj in df["Objektname"]]
                    
                    st.write("### 📋 Daten-Vorschau inklusive Varianten-Spalten:")
                    st.dataframe(df)
                    
                    # ==========================================
                    # 5. EXCEL EXPORT
                    # ==========================================
                    buffer_export = io.BytesIO()
                    with pd.ExcelWriter(buffer_export, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Objekt_Auswahl")
                        metriken_df.to_excel(writer, index=False, sheet_name="Varianten_Vergleich")
                        
                    st.download_button(
                        label="📥 Excel mit berechneten Varianten herunterladen",
                        data=buffer_export.getvalue(),
                        file_name="4d_optimierung_top10_schnell.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
    except Exception as e:
        st.error(f"Fehler bei der Verarbeitung: {e}")
