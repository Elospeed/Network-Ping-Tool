# Ideen

Sammlung von Ideen für die Weiterentwicklung. Neue Einträge oben anfügen, umgesetzte mit Version/PR abhaken.

## Offen

- [ ] **Lizenzfreie Variante als Lazarus-Anwendung** – Portable .exe ohne Excel-Lizenz (Free Pascal/Lazarus). Ping über `IcmpSendEcho` wie im VBA-Code, Fenster mit Ergebnisliste statt UserForm, IP-Liste mit Namen aus CSV, optional paralleles Pingen per Threads. LibreOffice/OpenOffice Portable verworfen (VBA-`Declare` mit Struct und UserForms nicht zuverlässig).
- [ ] **.0 und .255 im Range-Modus optional überspringen** – z. B. per Checkbox, damit Broadcast-Antworten nicht als «ONLINE» erscheinen (siehe [PROBLEME.md](PROBLEME.md)).
- [ ] **Hostnamen auflösen** – Statt «ungültig» den Namen per DNS auflösen (z. B. `getaddrinfo`) und dann pingen.

## Umgesetzt

- [x] **Gerätename in den Ergebnissen** (IP | Name | Status | RTT) – v1.1.0 / PR #1
- [x] **Stopp-Button**, der nur den Scan abbricht, ohne Excel zu schliessen – v1.1.0 / PR #1
- [x] **`<1 ms` statt `0 ms`** wie bei ping.exe – v1.1.0 / PR #1
