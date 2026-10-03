'==============================================================================
'  NETWORK PING TOOL - Excel VBA
'==============================================================================
'  Zweck:     Schnelles Pingen einzelner IPs, IP-Listen oder ganzer Ranges
'             im industriellen Netzwerk (PLCs, HMIs, Frequenzumrichter)
'
'  Technik:   - Windows ICMP API (iphlpapi.dll) für native Pings
'             - Array-Batch-Update für schnelle ListBox-Darstellung
'             - 32/64-Bit kompatible Declares
'
'  Autor:     T.Stöckli
'  Stand:     Juni 2026
'==============================================================================

Option Explicit

'==============================================================================
'  API-DEKLARATIONEN
'------------------------------------------------------------------------------
'  Diese Funktionen werden direkt aus Windows-DLLs aufgerufen.
'  Die VBA7-Variante (PtrSafe + LongPtr) ist für 64-Bit-Office nötig.
'==============================================================================

#If VBA7 Then
    ' ----- ICMP-PING-FUNKTIONEN (iphlpapi.dll) -----
    
    ' Erstellt einen Handle für ICMP-Operationen (1x am Anfang nötig)
    Private Declare PtrSafe Function IcmpCreateFile Lib "iphlpapi.dll" () As LongPtr
    
    ' Schließt den ICMP-Handle wieder (1x am Ende - WICHTIG für Speicherfreigabe!)
    Private Declare PtrSafe Function IcmpCloseHandle Lib "iphlpapi.dll" _
        (ByVal IcmpHandle As LongPtr) As Long
    
    ' Sendet einen Ping und wartet auf Antwort (das ist die Hauptfunktion!)
    Private Declare PtrSafe Function IcmpSendEcho Lib "iphlpapi.dll" ( _
        ByVal IcmpHandle As LongPtr, _
        ByVal DestinationAddress As Long, _
        ByVal RequestData As String, _
        ByVal RequestSize As Integer, _
        ByVal RequestOptions As Long, _
        ByRef ReplyBuffer As Byte, _
        ByVal ReplySize As Long, _
        ByVal Timeout As Long _
    ) As Long
    
    ' Konvertiert einen IP-String ("192.168.0.1") in einen Long-Wert
    Private Declare PtrSafe Function inet_addr Lib "wsock32.dll" _
        (ByVal cp As String) As Long
    
    ' Kopiert Speicherbereiche (für Reply-Buffer-Auswertung)
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    
    ' ----- KEYBOARD-SIMULATION für Screenshot (user32.dll) -----
    
    ' Simuliert Tastendruck (wird für Alt+PrintScreen genutzt)
    Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, _
         ByVal dwExtraInfo As LongPtr)
#Else
    ' ----- ALTE 32-BIT VARIANTEN (für sehr altes Office) -----
    Private Declare Function IcmpCreateFile Lib "iphlpapi.dll" () As Long
    Private Declare Function IcmpCloseHandle Lib "iphlpapi.dll" _
        (ByVal IcmpHandle As Long) As Long
    Private Declare Function IcmpSendEcho Lib "iphlpapi.dll" ( _
        ByVal IcmpHandle As Long, _
        ByVal DestinationAddress As Long, _
        ByVal RequestData As String, _
        ByVal RequestSize As Integer, _
        ByVal RequestOptions As Long, _
        ByRef ReplyBuffer As Byte, _
        ByVal ReplySize As Long, _
        ByVal Timeout As Long _
    ) As Long
    Private Declare Function inet_addr Lib "wsock32.dll" _
        (ByVal cp As String) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, _
         ByVal dwExtraInfo As Long)
#End If

'==============================================================================
'  KONSTANTEN
'==============================================================================

' Wie lange wartet ein Ping maximal auf Antwort? (in Millisekunden)
' - Im LAN: 30ms reicht völlig
' - Bei WLAN: ggf. höher (50-100ms)
Private Const PING_TIMEOUT As Long = 100

' Virtual-Key-Codes für Screenshot-Funktion
Private Const VK_SNAPSHOT As Long = &H2C      ' Print-Screen-Taste
Private Const VK_MENU As Long = &H12          ' Alt-Taste
Private Const KEYEVENTF_KEYUP As Long = &H2   ' Flag "Taste loslassen"

'==============================================================================
'  MODUL-VARIABLEN (für Scan-Status)
'==============================================================================

' Flag: Läuft gerade ein Scan?
' Wichtig für sauberes Schließen der Form während eines Scans
Private m_ScanRunning As Boolean

' Flag: Wurde der Scan abgebrochen (Stopp-Button oder Form wird geschlossen)?
Private m_ScanAbort As Boolean

' Flag: Wird die Form gerade geschlossen? (dann UI nach dem Scan nicht mehr antasten)
Private m_FormClosing As Boolean


'==============================================================================
'  HAUPT-SCAN-FUNKTION
'------------------------------------------------------------------------------
'  Pingt eine oder mehrere IP-Adressen und zeigt die Ergebnisse in der ListBox
'  (IP | Name | Status | RTT). Der Name kommt aus Spalte B von "IP-Liste".
'
'  Parameter:
'    targetIP - Die zu pingende IP (bei Range: nur das Subnet, z.B. 10.0.0.0)
'    isRange  - True = ganzer Range (.0-.255) wird gescannt
'    isList   - True = IP-Liste aus dem Tabellenblatt "IP-Liste" wird genutzt
'    (beide False = nur die targetIP wird gepingt)
'==============================================================================

Public Sub RunNetworkScan(ByVal targetIP As String, ByVal isRange As Boolean, _
                          ByVal isList As Boolean)
    
    ' ----- Variablen-Deklarationen -----
    Dim ipList As New Collection      ' Sammlung aller zu pingenden IPs
    Dim nameLookup As Collection      ' Gerätenamen aus "IP-Liste" (Key = normalisierte IP)
    Dim currentIP As String           ' Die gerade aktuell gepingte IP (normalisiert)
    Dim pingResult As String          ' Ergebnis vom Ping ("Offline" oder RTT-Wert)
    Dim i As Long, j As Long, k As Long  ' Schleifenzähler
    Dim scanned As Long               ' Anzahl tatsächlich gepingter IPs (bei Stopp < Gesamt)
    Dim ws As Worksheet               ' Referenz auf das IP-Liste-Tabellenblatt
    Dim lastRow As Long               ' Letzte beschriebene Zeile in der Tabelle
    Dim item As Variant               ' For-Each-Iteration über die Collection
    Dim startTime As Double           ' Zeitstempel für Performance-Messung
    Dim resultsArray() As Variant     ' 2D-Array mit allen Scan-Ergebnissen
    Dim displayArray() As Variant     ' Was am Ende in der ListBox landet (ggf. sortiert)
    Dim maxBarWidth As Double         ' Maximale Breite des Fortschrittsbalkens
    Dim nameWidth As Double           ' Breite der Namensspalte in der ListBox
    Dim errNum As Long, errDesc As String  ' Fehlerinfo für die Fehlerbehandlung
    
    ' Handle für ICMP-Operationen (LongPtr in 64-Bit, Long in 32-Bit)
    #If VBA7 Then
        Dim hIcmp As LongPtr
    #Else
        Dim hIcmp As Long
    #End If
    
    '--------------------------------------------------------------------------
    '  SCHRITT 1: IP-Adressen sammeln
    '--------------------------------------------------------------------------
    
    On Error Resume Next
    
    If isRange Then
        ' Range-Modus: Alle 256 IPs eines Subnets generieren (z.B. 10.0.0.0-255)
        Dim prefix As String
        prefix = Left(targetIP, InStrRev(targetIP, "."))  ' Z.B. "10.0.0."
        
        For i = 0 To 255
            ipList.Add prefix & i
        Next i
        
    ElseIf isList Then
        ' Listen-Modus: IPs aus dem Tabellenblatt "IP-Liste" laden (ab Zeile 2)
        Set ws = ThisWorkbook.Worksheets("IP-Liste")
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> "" Then
                ipList.Add CStr(ws.Cells(i, 1).Value)
            End If
        Next i
        
    Else
        ' Single-Modus: Nur eine einzige IP
        If targetIP <> "" Then ipList.Add targetIP
    End If
    
    ' Ab hier: Laufzeitfehler -> ScanFehler (Handle schliessen, Buttons freigeben)
    On Error GoTo ScanFehler
    
    ' Abbruch wenn keine IPs gefunden wurden
    If ipList.count = 0 Then
        MsgBox "Keine IP-Adressen gefunden!", vbExclamation
        Exit Sub
    End If
    
    ' Gerätenamen aus dem Tabellenblatt laden (Namensspalte in allen Modi)
    Set nameLookup = LoadNameLookup()
    
    '--------------------------------------------------------------------------
    '  SCHRITT 2: UI und Datenstrukturen vorbereiten
    '--------------------------------------------------------------------------
    
    ' Fortschrittsbalken zurücksetzen
    maxBarWidth = Me.lblBackground.Width
    Me.lblbar.Width = 0
    
    ' ListBox vorbereiten (4 Spalten: IP | Name | Status | RTT)
    With Me.lstResults
        .Clear
        .ColumnCount = 4
        ' Namensspalte bekommt die restliche Breite (mind. 90, sonst Scrollbalken)
        nameWidth = .Width - 80 - 50 - 40 - 15
        If nameWidth < 90 Then nameWidth = 90
        .ColumnWidths = "80;" & nameWidth & ";50;40"
    End With
    
    ' Ergebnis-Array erstellen (so groß wie die Anzahl IPs)
    ' Das Array wird in EINEM Schritt am Ende in die ListBox geladen
    ' (schneller als jeden Eintrag einzeln per .AddItem hinzuzufügen!)
    ReDim resultsArray(1 To ipList.count, 1 To 4)
    
    ' ICMP-Handle erstellen (1x für ALLE Pings - viel schneller als pro Ping!)
    hIcmp = IcmpCreateFile()
    If hIcmp = 0 Then
        MsgBox "Fehler: ICMP-Handle konnte nicht erstellt werden.", vbCritical
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------
    '  SCHRITT 3: Eigentlicher Scan-Durchlauf
    '--------------------------------------------------------------------------
    
    startTime = Timer                        ' Performance-Messung starten
    Application.ScreenUpdating = False       ' UI-Rendering pausieren (schneller!)
    i = 0                                    ' Counter zurücksetzen (WICHTIG!)
    
    ' Scan-Status setzen (für Stopp-Button und sauberes Schließen der Form)
    m_ScanRunning = True
    m_ScanAbort = False
    
    ' Buttons während Scan sperren (verhindert Doppel-Klicks und Chaos)
    Call SetScanButtons(True)
    
    For Each item In ipList
        ' Stopp-Button gedrückt oder Form geschlossen -> Scan sauber abbrechen
        If m_ScanAbort Then Exit For
        
        i = i + 1
        
        ' Führende Nullen entfernen: inet_addr liest "010" sonst als Oktal (= 8)!
        currentIP = NormalizeIP(CStr(item))
        
        If currentIP = "" Then
            ' Kein gültiges IPv4-Format (z.B. Hostname oder Tippfehler)
            resultsArray(i, 1) = CStr(item)
            resultsArray(i, 2) = ""
            resultsArray(i, 3) = "ungültig"
            resultsArray(i, 4) = "-"
        Else
            ' Eigentlicher Ping-Aufruf
            pingResult = FastPing(hIcmp, currentIP, PING_TIMEOUT)
        
            ' Ergebnis in Array schreiben
            resultsArray(i, 1) = currentIP
            resultsArray(i, 2) = LookupName(nameLookup, currentIP)
            
            If pingResult = "Offline" Then
                resultsArray(i, 3) = "offline"
                resultsArray(i, 4) = "-"
            ElseIf pingResult = "0" Then
                ' Wie ping.exe: Antworten unter 1 ms als "<1 ms" anzeigen
                resultsArray(i, 3) = "[ONLINE]"
                resultsArray(i, 4) = "<1 ms"
            Else
                resultsArray(i, 3) = "[ONLINE]"
                resultsArray(i, 4) = pingResult & " ms"
            End If
        End If
        
        ' Fortschrittsbalken aktualisieren
        Me.lblbar.Width = (i / ipList.count) * maxBarWidth
        
        ' Fortschritt im Fenstertitel anzeigen (alle 5 IPs aktualisieren - schneller)
        If ((i Mod 5) = 0 Or i = ipList.count) And Not m_ScanAbort Then
            Me.Caption = "Scanne... " & i & " / " & ipList.count & " IPs"
        End If
        
        DoEvents   ' UI-Events verarbeiten (Form bleibt responsive, Stopp-Button)
    Next item
    
    scanned = i
    
    ' ICMP-Handle schließen (Speicher freigeben!)
    IcmpCloseHandle hIcmp
    hIcmp = 0
    
    ' Scan-Status zurücksetzen
    m_ScanRunning = False
    
    ' Form wird geschlossen -> UI nicht mehr antasten
    If m_FormClosing Then Exit Sub
    
    ' Buttons wieder freigeben
    Call SetScanButtons(False)
    Application.ScreenUpdating = True
    
    ' Sofort gestoppt, bevor die erste IP gepingt wurde
    If scanned = 0 Then
        Me.Caption = "Scan gestoppt."
        Exit Sub
    End If
    
    '--------------------------------------------------------------------------
    '  SCHRITT 4: Ergebnisse übernehmen, optional sortiert (Online nach oben)
    '--------------------------------------------------------------------------
    
    ' Nur die tatsächlich gepingten Einträge (bei Stopp weniger als geplant)
    ReDim displayArray(1 To scanned, 1 To 4)
    k = 0
    
    If Me.chkSort.Value = True Then
        ' Stabil sortieren: zuerst alle Online-, danach alle übrigen Einträge,
        ' jeweils in der ursprünglichen (numerischen) Reihenfolge
        For i = 1 To scanned
            If resultsArray(i, 3) = "[ONLINE]" Then
                k = k + 1
                For j = 1 To 4: displayArray(k, j) = resultsArray(i, j): Next j
            End If
        Next i
        For i = 1 To scanned
            If resultsArray(i, 3) <> "[ONLINE]" Then
                k = k + 1
                For j = 1 To 4: displayArray(k, j) = resultsArray(i, j): Next j
            End If
        Next i
    Else
        For i = 1 To scanned
            For j = 1 To 4: displayArray(i, j) = resultsArray(i, j): Next j
        Next i
    End If
    
    '--------------------------------------------------------------------------
    '  SCHRITT 5: Ergebnisse anzeigen
    '--------------------------------------------------------------------------
    
    ' Komplettes Array auf einen Schlag in die ListBox laden (sehr schnell!)
    Me.lstResults.List = displayArray
    
    ' Performance-Info im Fenstertitel anzeigen
    If m_ScanAbort Then
        Me.Caption = "Scan gestoppt! " & scanned & " von " & ipList.count & _
                     " IPs in " & Format(Timer - startTime, "0.0") & " Sek."
    Else
        Me.Caption = "Scan fertig! " & ipList.count & " IPs in " & _
                     Format(Timer - startTime, "0.0") & " Sek."
    End If
    Exit Sub
    
    '--------------------------------------------------------------------------
    '  FEHLERBEHANDLUNG: Laufzeitfehler mitten im Scan
    '  Handle schliessen und Buttons freigeben, damit die Form bedienbar bleibt
    '--------------------------------------------------------------------------
ScanFehler:
    errNum = Err.Number
    errDesc = Err.Description
    Resume ScanAufraeumen
    
ScanAufraeumen:
    On Error Resume Next
    If hIcmp <> 0 Then IcmpCloseHandle hIcmp
    m_ScanRunning = False
    Application.ScreenUpdating = True
    
    If Not m_FormClosing Then
        Call SetScanButtons(False)
        Me.Caption = "Scan mit Fehler beendet"
        MsgBox "Fehler beim Scan (" & errNum & "): " & errDesc, vbCritical
    End If
End Sub


'==============================================================================
'  FASTPING - Pingt eine einzelne IP via Windows ICMP API
'------------------------------------------------------------------------------
'  Parameter:
'    hIcmp    - ICMP-Handle (von IcmpCreateFile)
'    strIP    - IP-Adresse als String (z.B. "192.168.0.1")
'    Timeout  - Wartezeit auf Antwort in Millisekunden
'
'  Rückgabe:
'    "Offline"         - Wenn die IP nicht erreichbar ist
'    "<Zahl>"          - Round-Trip-Time in Millisekunden (z.B. "4")
'==============================================================================

#If VBA7 Then
Private Function FastPing(ByVal hIcmp As LongPtr, ByVal strIP As String, _
                          Optional ByVal Timeout As Long = 100) As String
#Else
Private Function FastPing(ByVal hIcmp As Long, ByVal strIP As String, _
                          Optional ByVal Timeout As Long = 100) As String
#End If
    Dim lAddr As Long              ' IP als Long-Zahl (für die API)
    Dim lReply As Long             ' Anzahl der empfangenen Antworten (0 = offline)
    Dim buffer(0 To 1023) As Byte  ' Empfangs-Buffer für die Reply-Daten
    Dim Status As Long             ' Status-Wert aus dem Reply (0 = Erfolg)
    Dim rtt As Long                ' Round-Trip-Time in ms
    
    ' Standardwert: Offline (wird überschrieben wenn Reply kommt)
    FastPing = "Offline"
    
    ' IP-String in Long konvertieren ("192.168.0.1" -> Zahl)
    lAddr = inet_addr(strIP)
    If lAddr = -1 Then Exit Function   ' Ungültige IP -> Abbruch
    
    ' Eigentlicher Ping-Aufruf an Windows
    lReply = IcmpSendEcho(hIcmp, lAddr, "PingData", 8, 0, _
                          buffer(0), 1024, Timeout)
    
    If lReply > 0 Then
        ' Antwort erhalten - jetzt Status und RTT aus dem Buffer extrahieren
        ' Buffer-Struktur (ICMP_ECHO_REPLY):
        '   Offset 0-3:  Address (IP der antwortenden Seite)
        '   Offset 4-7:  Status (0 = Erfolg)
        '   Offset 8-11: RoundTripTime (RTT in ms)
        
        CopyMemory Status, buffer(4), 4
        
        If Status = 0 Then
            CopyMemory rtt, buffer(8), 4
            FastPing = rtt   ' Rückgabe als String, z.B. "4"
        End If
    End If
End Function


Private Sub lblInfo_Click()
  FormInfo.Show
End Sub

'==============================================================================
'  TEXTBOX-LOGIK FÜR IP-EINGABE
'------------------------------------------------------------------------------
'  Vier TextBoxen für die vier IP-Oktette (z.B. 192 . 168 . 0 . 1)
'  Features:
'    - Auto-Hüpfen zur nächsten Box bei 3 Stellen
'    - Punkt-Eingabe springt sofort zur nächsten Box
'    - Maximum 255 wird erzwungen
'    - Nur Zahlen erlaubt (Punkte/Buchstaben werden entfernt)
'    - Beim Reinklicken/Tabben wird der Inhalt markiert
'==============================================================================

' ----- Change-Events: Werden bei jeder Eingabe ausgelöst -----
Private Sub TextBox1_Change(): Call ValidateIPBox(Me.TextBox1, Me.TextBox2): End Sub
Private Sub TextBox2_Change(): Call ValidateIPBox(Me.TextBox2, Me.TextBox3): End Sub
Private Sub TextBox3_Change(): Call ValidateIPBox(Me.TextBox3, Me.TextBox4): End Sub
Private Sub TextBox4_Change(): Call ValidateIPBox(Me.TextBox4, Nothing):      End Sub

'------------------------------------------------------------------------------
' Validiert eine IP-Box: nur GANZE Zahlen (0-255), Auto-Hüpfen bei 3 Stellen
'------------------------------------------------------------------------------
Private Sub ValidateIPBox(ByRef CurrentBox As MSForms.TextBox, _
                          ByRef NextBox As MSForms.Control)
    Dim cleaned As String
    Dim c As Long
    
    If CurrentBox.Text <> "" Then
        ' Nur Ziffern behalten (IsNumeric würde auch "2.2", "1-" oder "1e2" akzeptieren)
        For c = 1 To Len(CurrentBox.Text)
            If Mid$(CurrentBox.Text, c, 1) Like "[0-9]" Then
                cleaned = cleaned & Mid$(CurrentBox.Text, c, 1)
            End If
        Next c
        
        If cleaned <> CurrentBox.Text Then
            ' Löst Change erneut aus -> dort wird der bereinigte Text weiter geprüft
            CurrentBox.Text = cleaned
            Exit Sub
        End If
        
        ' Maximum 255 erzwingen
        If CLng(CurrentBox.Text) > 255 Then
            MsgBox "Maximal 255!", vbExclamation
            CurrentBox.Text = "255"
        End If
        
        ' Auto-Hüpfen zur nächsten Box wenn 3 Stellen erreicht
        If Len(CurrentBox.Text) = 3 And Not NextBox Is Nothing Then
            NextBox.SetFocus
            Call SelectAllText(NextBox)
        End If
    End If
End Sub

' ----- KeyPress-Events: Punkt-Eingabe springt zur nächsten Box -----
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then  ' 46 = ASCII für "."
        KeyAscii = 0       ' Punkt nicht in Box schreiben
        Me.TextBox2.SetFocus
        Call SelectAllText(Me.TextBox2)
    End If
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 0
        Me.TextBox3.SetFocus
        Call SelectAllText(Me.TextBox3)
    End If
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Then
        KeyAscii = 0
        Me.TextBox4.SetFocus
        Call SelectAllText(Me.TextBox4)
    End If
End Sub

' In TextBox4 gibt es keine nächste Box - Punkt einfach blockieren
Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 46 Or KeyAscii = 44 Then  ' Punkt oder Komma blockieren
        KeyAscii = 0
    End If
End Sub

' ----- Enter-Events: Inhalt beim Reinklicken/Tabben markieren -----
Private Sub TextBox1_Enter(): Call SelectAllText(Me.TextBox1): End Sub
Private Sub TextBox2_Enter(): Call SelectAllText(Me.TextBox2): End Sub
Private Sub TextBox3_Enter(): Call SelectAllText(Me.TextBox3): End Sub
Private Sub TextBox4_Enter(): Call SelectAllText(Me.TextBox4): End Sub

'------------------------------------------------------------------------------
' Hilfsfunktion: Markiert den kompletten Text einer Control
'------------------------------------------------------------------------------
Private Sub SelectAllText(ByRef ctrl As MSForms.Control)
    On Error Resume Next
    ctrl.SelStart = 0
    ctrl.SelLength = Len(ctrl.Text)
End Sub


'==============================================================================
'  BUTTON-EVENTS
'==============================================================================

'------------------------------------------------------------------------------
' SCAN-BUTTON: Startet den Scan basierend auf der gewählten Option
'   - obtSingle: Einzelne IP (aus den TextBoxen)
'   - obtList:   IP-Liste aus dem Tabellenblatt
'   - obtRange:  Kompletter Range .0-.255
'------------------------------------------------------------------------------
Private Sub btnScan_Click()
   Dim fullIP As String
    
   ' Scan Mode prüfen
   If obtSingle Or obtRange Then
      ' Prüfen ob alle 4 TextBoxen befüllt sind
      If Me.TextBox1.Text = "" Or Me.TextBox2.Text = "" Or _
         Me.TextBox3.Text = "" Or Me.TextBox4.Text = "" Then
        MsgBox "IP unvollständig!", vbExclamation
          Exit Sub
      End If
    End If
    
    ' IP aus den 4 TextBoxen zusammenbauen
    fullIP = TextBox1.Value & "." & TextBox2.Value & "." & _
             TextBox3.Value & "." & TextBox4.Value
    
    ' Je nach gewählter Option den richtigen Scan-Modus starten
    If obtSingle And Not TextBox1.Value = "" Then
        Call RunNetworkScan(fullIP, False, False)
    End If
    If obtList Then
        Call RunNetworkScan("", False, True)
    End If
    If obtRange Then
        Call RunNetworkScan(fullIP, True, False)
    End If
End Sub

'------------------------------------------------------------------------------
' STOPP-BUTTON: Bricht einen laufenden Scan ab, ohne Excel zu schliessen
'   Der Button "btnStop" muss im Formular-Designer angelegt werden. Fehlt er,
'   wird dieser Code einfach nie aufgerufen.
'   Bereits gepingte IPs bleiben in der Ergebnisliste stehen.
'------------------------------------------------------------------------------
Private Sub btnStop_Click()
    If m_ScanRunning Then
        m_ScanAbort = True
        Me.Caption = "Scan wird gestoppt..."
    End If
End Sub

'------------------------------------------------------------------------------
' Sperrt/entsperrt die Buttons während eines Scans
'   btnStop wird über Me.Controls angesprochen, damit der Code auch ohne
'   diesen Button kompiliert (dann wird er einfach übersprungen)
'------------------------------------------------------------------------------
Private Sub SetScanButtons(ByVal scanning As Boolean)
    Me.btnScan.Enabled = Not scanning
    Me.btnAddIP.Enabled = Not scanning
    Me.btnDeleteIP.Enabled = Not scanning
    
    On Error Resume Next
    Me.Controls("btnStop").Enabled = scanning
End Sub

'------------------------------------------------------------------------------
' ADD-IP-BUTTON: Fügt die aktuelle IP zur Liste hinzu (inkl. Gerätename)
'------------------------------------------------------------------------------
Private Sub btnAddIP_Click()
    Dim ws As Worksheet
    Dim newIP As String
    Dim deviceName As String
    Dim lastRow As Long
    
    ' Prüfen ob alle 4 TextBoxen befüllt sind
    If Me.TextBox1.Text = "" Or Me.TextBox2.Text = "" Or _
       Me.TextBox3.Text = "" Or Me.TextBox4.Text = "" Then
        MsgBox "IP unvollständig!", vbExclamation
        Exit Sub
    End If
    
    ' IP zusammenbauen (führende Nullen entfernen: "010" -> "10")
    newIP = NormalizeIP(Me.TextBox1.Text & "." & Me.TextBox2.Text & "." & _
                        Me.TextBox3.Text & "." & Me.TextBox4.Text)
    If newIP = "" Then
        MsgBox "IP ungültig!", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("IP-Liste")
    
    ' Prüfen ob die IP schon existiert (Duplikate verhindern)
    If Not ws.Columns("A").Find(newIP, LookAt:=xlWhole) Is Nothing Then
        MsgBox "IP existiert bereits!", vbCritical
        Exit Sub
    End If
    
    ' Gerätename abfragen
    deviceName = InputBox("Name für " & newIP & ":", "Neues Gerät")
    If Trim(deviceName) = "" Then Exit Sub   ' Abbruch bei leerer Eingabe
    
    ' In nächste freie Zeile schreiben
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = newIP
    ws.Cells(lastRow, 2).Value = deviceName
    
    ' Tabelle NUMERISCH nach IP sortieren
    Call SortIPTable(ws)
    
    Call UserForm_Initialize   ' IP-Liste neu laden
    
    ThisWorkbook.Save 'Änderungen speichern
    
    MsgBox "Gespeichert!", vbInformation
    
    ' Eingabe-TextBoxen leeren und Fokus zurück auf erste Box
    Me.TextBox1.Text = ""
    Me.TextBox2.Text = ""
    Me.TextBox3.Text = ""
    Me.TextBox4.Text = ""
    Me.TextBox1.SetFocus
End Sub

'------------------------------------------------------------------------------
' SCREENSHOT-BUTTON: Macht ein Bildschirmfoto des aktiven Fensters
'   Simuliert Alt+PrintScreen -> Bild liegt dann in der Zwischenablage
'------------------------------------------------------------------------------
Private Sub btnScreenshot_Click()
    DoEvents
    
    ' Alt-Taste drücken
    keybd_event VK_MENU, 0, 0, 0
    
    ' Print-Screen drücken und sofort wieder loslassen
    keybd_event VK_SNAPSHOT, 0, 0, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_KEYUP, 0
    
    ' Alt-Taste wieder loslassen
    keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
    
    DoEvents
    MsgBox "Screenshot in Zwischenablage!", vbInformation
End Sub

'------------------------------------------------------------------------------
' DELETE-IP-BUTTON: Entfernt den ausgewählten Eintrag aus der Liste
'------------------------------------------------------------------------------
Private Sub btnDeleteIP_Click()
    If lstList.ListIndex = -1 Then
        MsgBox "Eintrag wählen!", vbExclamation
        Exit Sub
    End If
    
    Me.lstList.RemoveItem Me.lstList.ListIndex
    Call SaveListBoxToWorksheet
    ThisWorkbook.Save 'Änderungen speichern
    
End Sub

'------------------------------------------------------------------------------
' Schreibt den aktuellen Inhalt der lstList zurück ins Tabellenblatt
' (wird nach jedem Löschen aufgerufen, damit Tabelle synchron bleibt)
'------------------------------------------------------------------------------
Private Sub SaveListBoxToWorksheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim lr As Long
    
    Set ws = ThisWorkbook.Worksheets("IP-Liste")
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    ' Alte Daten löschen (ab Zeile 2, Header bleibt erhalten)
    If lr >= 2 Then ws.Range("A2:B" & lr).ClearContents
    
    ' Neue Daten aus der ListBox schreiben
    For i = 0 To Me.lstList.ListCount - 1
        ws.Cells(i + 2, 1).Value = Me.lstList.List(i, 0)   ' Spalte A: IP
        ws.Cells(i + 2, 2).Value = Me.lstList.List(i, 1)   ' Spalte B: Name
    Next i
    
    Call SortIPTable(ws)
End Sub


'==============================================================================
'  DOPPELKLICK-HANDLER
'------------------------------------------------------------------------------
'  Bei DblClick auf einen Eintrag wird die IP automatisch in die TextBoxen
'  übernommen - praktisch für schnelle Einzeltests.
'==============================================================================

Private Sub lstList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call FillIPBoxes(Me.lstList)
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call FillIPBoxes(Me.lstResults)
End Sub

'------------------------------------------------------------------------------
' Splittet die IP aus der ListBox auf und füllt die 4 TextBoxen
'------------------------------------------------------------------------------
Private Sub FillIPBoxes(ByRef lst As MSForms.ListBox)
    Dim ipParts() As String
    
    If lst.ListIndex = -1 Then Exit Sub
    
    ' "192.168.0.1" -> Array("192", "168", "0", "1")
    ipParts = Split(lst.List(lst.ListIndex, 0), ".")
    
    ' Nur wenn wirklich 4 Teile entstanden sind
    If UBound(ipParts) = 3 Then
        Me.TextBox1.Value = ipParts(0)
        Me.TextBox2.Value = ipParts(1)
        Me.TextBox3.Value = ipParts(2)
        Me.TextBox4.Value = ipParts(3)
    End If
End Sub


'==============================================================================
'  USERFORM QUERY CLOSE
'------------------------------------------------------------------------------
'  Wird beim Schließen der Form ausgelöst (X-Button, Esc, etc.)
'  Verhindert Absturz wenn während eines Scans geschlossen wird.
'==============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Wenn gerade ein Scan läuft: Abbruch signalisieren
    ' Warten bringt hier nichts: Die Scan-Schleife steckt im Call-Stack UNTER
    ' diesem Event (in DoEvents) und läuft erst weiter, wenn dieses Event fertig
    ' ist. Dann erkennt sie m_ScanAbort, schliesst das ICMP-Handle und fasst
    ' wegen m_FormClosing die Form nicht mehr an.
    m_FormClosing = True
    If m_ScanRunning Then m_ScanAbort = True
    
    ' Excel sauber schließen
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    
    Dim wb As Workbook
    Dim otherWbOpen As Boolean
    otherWbOpen = False
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            otherWbOpen = True
            Exit For
        End If
    Next wb
    
    If otherWbOpen Then
        ThisWorkbook.Close SaveChanges:=False
    Else
        Application.Quit
        ThisWorkbook.Close SaveChanges:=False
    End If
    

End Sub


'==============================================================================
'  USERFORM INITIALIZE
'------------------------------------------------------------------------------
'  Wird beim Öffnen der Form ausgeführt: Lädt die IP-Liste aus dem Tabellenblatt
'  in die rechte ListBox (lstList).
'==============================================================================

Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim i As Long
    Dim lr As Long
    
    ' IP-TextBoxen auf maximal 3 Zeichen begrenzen (verhindert "1234" etc.)
    Me.TextBox1.MaxLength = 3
    Me.TextBox2.MaxLength = 3
    Me.TextBox3.MaxLength = 3
    Me.TextBox4.MaxLength = 3
    
    ' Stopp-Button (optional) ist nur während eines Scans aktiv
    On Error Resume Next
    Me.Controls("btnStop").Enabled = False
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Worksheets("IP-Liste")
    
    ' ListBox einrichten (2 Spalten: IP | Name)
    With Me.lstList
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "90;" & (Me.lstList.Width - 80)
    End With
    
    ' Daten aus Tabelle laden (ab Zeile 2 - Zeile 1 ist Header)
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lr >= 2 Then
        For i = 2 To lr
            Me.lstList.AddItem ws.Cells(i, 1).Value
            Me.lstList.List(Me.lstList.ListCount - 1, 1) = ws.Cells(i, 2).Value
        Next i
    End If
End Sub



'------------------------------------------------------------------------------
' Sortiert die IP-Tabelle NUMERISCH (nach Oktetten, nicht alphabetisch!)
'   Beispiel: 10.0.0.2 kommt VOR 10.0.0.10 (anders als bei Text-Sort!)
'   Leerzeilen, Hostnamen und Tippfehler landen unten (in bisheriger
'   Reihenfolge), statt die Sortierung abstürzen zu lassen.
'------------------------------------------------------------------------------

Private Sub SortIPTable(ByRef ws As Worksheet)
    ' Schlüssel für ungültige Einträge: grösser als jede IP (max. 2^32 - 1)
    Const INVALID_KEY As Double = 4294967296#
    
    Dim lastRow As Long
    Dim data As Variant            ' Spalten A:B als Array (schneller als Zelle für Zelle)
    Dim keys() As Double           ' Numerischer Sortierschlüssel pro Zeile
    Dim i As Long, j As Long, n As Long
    Dim tmpKey As Double, tmpIP As Variant, tmpName As Variant
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow < 3 Then Exit Sub   ' Weniger als 2 Einträge - nichts zu sortieren
    
    data = ws.Range("A2:B" & lastRow).Value
    n = UBound(data, 1)
    ReDim keys(1 To n)
    
    For i = 1 To n
        If IsError(data(i, 1)) Then
            keys(i) = INVALID_KEY
        Else
            keys(i) = IPSortKey(CStr(data(i, 1)))
            If keys(i) < 0 Then keys(i) = INVALID_KEY
        End If
    Next i
    
    ' Insertion-Sort (stabil: gleiche Schlüssel behalten ihre Reihenfolge)
    ' Bei <100 IPs schnell genug, einfach und verständlich
    For i = 2 To n
        tmpKey = keys(i)
        tmpIP = data(i, 1)
        tmpName = data(i, 2)
        
        j = i - 1
        Do While j >= 1
            If keys(j) <= tmpKey Then Exit Do
            keys(j + 1) = keys(j)
            data(j + 1, 1) = data(j, 1)
            data(j + 1, 2) = data(j, 2)
            j = j - 1
        Loop
        
        keys(j + 1) = tmpKey
        data(j + 1, 1) = tmpIP
        data(j + 1, 2) = tmpName
    Next i
    
    ' Vorher leeren, damit leere Zellen im Array auch wirklich leer geschrieben werden
    ws.Range("A2:B" & lastRow).ClearContents
    ws.Range("A2:B" & lastRow).Value = data
End Sub

'------------------------------------------------------------------------------
' Numerischer Sortierschlüssel einer IP (10.0.0.2 < 10.0.0.10)
'   Rückgabe -1 wenn kein gültiges IPv4-Format
'------------------------------------------------------------------------------
Private Function IPSortKey(ByVal strIP As String) As Double
    Dim parts() As String
    
    strIP = NormalizeIP(strIP)
    If strIP = "" Then
        IPSortKey = -1
        Exit Function
    End If
    
    parts = Split(strIP, ".")
    IPSortKey = ((CDbl(parts(0)) * 256 + CDbl(parts(1))) * 256 + _
                 CDbl(parts(2))) * 256 + CDbl(parts(3))
End Function

'------------------------------------------------------------------------------
' Normalisiert eine IPv4-Adresse: "192.168.001.010" -> "192.168.1.10"
'   WICHTIG: inet_addr liest Oktette mit führender 0 als Oktalzahl
'   ("010" = 8!). Deshalb wird jedes Oktett vorher dezimal umgewandelt.
'   Rückgabe "" wenn kein gültiges IPv4-Format (Hostname, Tippfehler, leer)
'------------------------------------------------------------------------------
Private Function NormalizeIP(ByVal strIP As String) As String
    Dim parts() As String
    Dim p As Long, c As Long
    
    parts = Split(Trim(strIP), ".")
    If UBound(parts) <> 3 Then Exit Function   ' Genau 4 Teile nötig
    
    For p = 0 To 3
        ' Nur 1-3 Ziffern pro Oktett erlaubt
        If Len(parts(p)) = 0 Or Len(parts(p)) > 3 Then Exit Function
        For c = 1 To Len(parts(p))
            If Not (Mid$(parts(p), c, 1) Like "[0-9]") Then Exit Function
        Next c
        
        If CLng(parts(p)) > 255 Then Exit Function
        parts(p) = CStr(CLng(parts(p)))   ' "010" -> "10"
    Next p
    
    NormalizeIP = Join(parts, ".")
End Function

'------------------------------------------------------------------------------
' Liest die Gerätenamen aus "IP-Liste" (Spalte A = IP, Spalte B = Name)
'   Key = normalisierte IP, damit z.B. "192.168.010.5" auch gefunden wird
'------------------------------------------------------------------------------
Private Function LoadNameLookup() As Collection
    Dim ws As Worksheet
    Dim lr As Long, i As Long
    Dim ip As String
    Dim lookup As New Collection
    
    Set LoadNameLookup = lookup
    
    ' Blatt fehlt, doppelte IP (erster Name gewinnt) oder Fehlerwerte -> überspringen
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("IP-Liste")
    If ws Is Nothing Then Exit Function
    
    lr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For i = 2 To lr
        ip = ""
        ip = NormalizeIP(CStr(ws.Cells(i, 1).Value))
        If ip <> "" Then lookup.Add CStr(ws.Cells(i, 2).Value), ip
    Next i
End Function

'------------------------------------------------------------------------------
' Gerätename zu einer (normalisierten) IP, "" wenn nicht in der Liste
'------------------------------------------------------------------------------
Private Function LookupName(ByVal lookup As Collection, ByVal ip As String) As String
    On Error Resume Next
    LookupName = lookup(ip)
End Function

Public Sub SpendenLink_Oeffnen()
    ' Öffnet den oben definierten Link im Standardbrowser des Nutzers.
    ' ThisWorkbook.FollowHyperlink funktioniert zuverlässig unter Windows
    ' und macOS, ohne dass Shell/WScript benötigt wird.
    On Error GoTo Fehlerbehandlung
 
    ThisWorkbook.FollowHyperlink Address:="https://ko-fi.com/Elospeed", NewWindow:=True
 
    Exit Sub
 
Fehlerbehandlung:
    MsgBox "Der Link konnte nicht geöffnet werden." & vbCrLf & _
           "Bitte besuche manuell: " & "https://ko-fi.com/Elospeed", _
           vbExclamation, "Fehler beim Öffnen des Links"
End Sub


