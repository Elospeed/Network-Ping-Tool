# Probleme

Bekannte Fehler und Schwachstellen. Neue Einträge oben anfügen, erledigte mit Datum und Commit/PR abhaken.

## Offen

- [ ] **Range-Modus pingt .0 und .255** – Auf `.255` kann ein Broadcast-Reply kommen, der als «ONLINE» erscheint. In Subnetzen grösser als /24 sind .0/.255 aber gültige Hosts, deshalb noch nicht geändert.
- [ ] **Neue Version noch nicht in der .xlsm** – `quellcode.bas` ist korrigiert, muss aber noch ins Formular `FormPingTool` der `Network-Ping-Tool.xlsm` übernommen werden (inkl. Button `btnStop`, `lstResults` verbreitern).
- [ ] **Autor im Code-Header** – Im öffentlichen Header steht noch der Klarname. Ändern, falls öffentlich nur «Elospeed» erscheinen soll. Ältere Commits laufen noch unter dem früheren Usernamen.

## Erledigt

- [x] **Führende Nullen wurden als Oktal gelesen** – `192.168.001.010` pingte in Wirklichkeit `192.168.1.8`. Jede IP wird jetzt vor dem Ping normalisiert. (PR #1)
- [x] **SortIPTable stürzte ab** bei Leerzeilen, Hostnamen oder Tippfehlern im Blatt «IP-Liste» (Fehler 9 / 13). Ungültige Zeilen landen jetzt unten. (PR #1)
- [x] **README stimmte nicht**: Dateiname, Hostnamen, Abbruch. (PR #1)
- [x] **«Online nach oben» brachte Offline-Einträge durcheinander** – Sortierung ist jetzt stabil. (PR #1)
- [x] **Wirkungslose Warteschleife in QueryClose** entfernt. (PR #1)
- [x] **Kein Error-Handler in RunNetworkScan** – Buttons blieben nach einem Laufzeitfehler gesperrt. (PR #1)
- [x] **IsNumeric akzeptierte «1-»** – reiner Ziffernfilter in den IP-Feldern. (PR #1)
