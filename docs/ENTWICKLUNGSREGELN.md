# Entwicklungsregeln – Network Ping Tool

Verbindliche Regeln für alle, die an diesem Repository arbeiten (Menschen und Werkzeuge). **Vor jeder Arbeit an der Oberfläche (UserForm, Steuerelemente, Layout) diese Datei lesen.**

Die Regeln gelten gleich für alle Elospeed-Programme (StemMaker mit allen AddOns, WinKeeper, LedPreResistor, Network Ping Tool und jedes neue Programm). Die Lazarus-Programme haben dazu eine gemeinsame Skalierungs-Unit; hier im Excel-Werkzeug gilt dasselbe Ziel mit den Mitteln von VBA.

## 1. Windows-Skalierung (DPI) – gilt für jedes Fenster

**Warum:** Im Oktober 2026 waren mehrere Elospeed-Programme bei 125–200 % Windows-Skalierung kaputt: Fenster höher als der Bildschirm, Mindestgrösse gleich Startgrösse, Inhalte überlappten. Gemeldet hat es ein Benutzer, nicht unser Test. Das darf nicht nochmals passieren.

**Ziel:** Die UserForm und jede Meldung sehen bei **100 %, 125 %, 150 % und 200 %** richtig aus.

### Regeln

1. **Grössen in Punkten denken, nicht in Pixeln.** VBA-UserForms rechnen `Width`, `Height`, `Left`, `Top` in Punkten (1 pt = 1/72 Zoll); Windows skaliert sie mit. Keine Werte aus Pixel-APIs (z. B. `GetSystemMetrics`) ohne Umrechnung mit der aktuellen DPI übernehmen.
2. **Nichts doppelt skalieren:** Werte, die schon in Punkten vorliegen, nicht nochmals mit dem Skalierungsfaktor multiplizieren.
3. **Nie grösser als der freie Bildschirm:** Beim Öffnen (`UserForm_Initialize`/`Activate`) Höhe und Breite auf `Application.UsableHeight`/`UsableWidth` begrenzen; reicht der Platz nicht, Liste und Inhalt verkleinern statt Knöpfe abzuschneiden.
4. **Abhängige Grössen aus dem Formular ableiten** (z. B. Fortschrittsbalken aus `lblBackground.Width`), nicht als feste Zahl.
5. **Spaltenbreiten der ListBox** (`ColumnWidths`) so wählen, dass bei 200 % nichts abgeschnitten wird.
6. **Keine Maximieren-Funktion**, solange das Layout fest ist.

### Test – vor jedem PR mit GUI-Änderung und vor jedem Release

- Unter Windows (oder Excel unter Wine) mit **100 % und 150 %** testen, bei Platzproblemen auch **200 %** (Full HD). Screenshots zur Kontrolle.
- Excel nach dem Ändern der Skalierung neu starten.
- Im PR schreiben, bei welchen Stufen getestet wurde.
