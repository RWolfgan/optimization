# 4D Zielwert-Optimierer

## Browser-Version für Kollegen

Die Datei `index.html` verarbeitet Excel-Dateien und berechnet die Varianten vollständig im Browser. Anzahl der Varianten und maximale Rechenzeit je Variante können in der Oberfläche eingestellt werden. Es werden keine Nutzdaten an einen Server übertragen. Beim Öffnen werden lediglich die JavaScript-Bibliotheken SheetJS und GLPK.js über ein CDN geladen.

Zum lokalen Testen im Projektordner:

```bash
python3 -m http.server 8000
```

Anschließend `http://localhost:8000` öffnen. Für Kollegen kann die statische Seite beispielsweise über GitHub Pages, einen internen Webserver oder einen freigegebenen Ordner bereitgestellt werden. Beim direkten Doppelklick auf `index.html` funktioniert sie in modernen Browsern meist ebenfalls; ein Webserver ist zuverlässiger.

## Bestehende Streamlit-Version

```bash
pip install -r requirements.txt
streamlit run app.py
```
