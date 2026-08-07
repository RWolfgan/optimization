# 4D Zielwert-Optimierer

## Streamlit-Apps

Die aktuelle `app.py` ordnet Konten aus einer Summen- und Saldenliste höchstens einer frei definierten BWA-Position zu; nicht passende Konten dürfen unzugeordnet bleiben. Die Eingabe erfolgt über eine Excel-Datei mit den Tabellenblättern `Konten` und `BWA`; beide enthalten dieselben Jahre als Spalten. Die App minimiert mit CBC die absolute Abweichung aller BWA-Positionen und Jahre gemeinsam.

```bash
streamlit run app.py
```

Die bisherige 4D-Kombinations-App bleibt als `app_4d_optimizer.py` erhalten:

```bash
streamlit run app_4d_optimizer.py
```

## Browser-Version für Kollegen

Die Datei `index.html` verarbeitet Excel-Dateien und berechnet die Varianten vollständig im Browser. Anzahl der Varianten und maximale Rechenzeit je Variante können in der Oberfläche eingestellt werden. Die herunterladbare Vorlage enthält außerdem das Blatt `Zielwerte`; dort eingetragene Werte werden beim Auswählen der Datei automatisch in die Oberfläche übernommen. Es werden keine Nutzdaten an einen Server übertragen. Beim Öffnen werden lediglich die JavaScript-Bibliotheken SheetJS und GLPK.js über ein CDN geladen.

Zum lokalen Testen im Projektordner:

```bash
python3 -m http.server 8000
```

Anschließend `http://localhost:8000` öffnen. Für Kollegen kann die statische Seite beispielsweise über GitHub Pages, einen internen Webserver oder einen freigegebenen Ordner bereitgestellt werden. Beim direkten Doppelklick auf `index.html` funktioniert sie in modernen Browsern meist ebenfalls; ein Webserver ist zuverlässiger.

Die benötigten Python-Pakete werden mit `pip install -r requirements.txt` installiert.
