# Ideen

Sammlung von Ideen für die Weiterentwicklung. Neue Einträge oben anfügen, umgesetzte mit Version/PR abhaken.

## Offen

- [ ] **Lizenzfreie Variante ohne Excel** – LibreOffice/OpenOffice Portable reicht nicht ohne Umbau (OpenOffice kaum VBA; LibreOffice: `Declare IcmpSendEcho` mit Struct unzuverlässig, UserForm wird nur grob konvertiert). Optionen: PowerShell-Skript mit `Test-Connection` und CSV-IP-Liste (empfohlen) oder LibreOffice-Makro mit WMI `Win32_PingStatus` und eigenem Dialog.
- [ ] **.0 und .255 im Range-Modus optional überspringen** – z. B. per Checkbox, damit Broadcast-Antworten nicht als «ONLINE» erscheinen (siehe [PROBLEME.md](PROBLEME.md)).
- [ ] **Hostnamen auflösen** – Statt «ungültig» den Namen per DNS auflösen (z. B. `getaddrinfo`) und dann pingen.

## Umgesetzt

- [x] **Gerätename in den Ergebnissen** (IP | Name | Status | RTT) – v1.1.0 / PR #1
- [x] **Stopp-Button**, der nur den Scan abbricht, ohne Excel zu schliessen – v1.1.0 / PR #1
- [x] **`<1 ms` statt `0 ms`** wie bei ping.exe – v1.1.0 / PR #1
