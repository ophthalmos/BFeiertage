;   Funktion ......: Osterberechnung
;   Urheber .......: Dr. Heiner Lichtenberg
;                    Abwandlung der Gaußschen Osterformel zur Beseitigung der Ausnahmeregeln
;                    HISTORIA MATHEMATICA 24 (1997), 441–444 // ARTICLE NO. HM972170
;                    Quelle: http://www.sciencedirect.com/science/article/pii/S0315086097921704
;   AHK Version ...: AHK_L 1.1.13.00 x64 Unicode
;   Win Version ...: Windows 7 Ultimate x64 SP1
;   Author(s) .....: Seidenweber

; =='GLOBALE EINSTELLUNGEN'======================================================================================================

    #NoEnv                      ; Nicht nachsehen, ob leere Varibalen evtl. Umgebungsvariablen sind
    #SingleInstance force       ; Bei Neustart des Scriptes die alte Instanz ohne Nachfrage ersetzen
    SetBatchLines -1            ; Das Script läuft ohne Zwangspausen. (nur für schnelle Rechner empfohlen)
    SetControlDelay, -1         ; Wartezeit beim Zugriff auf langsame Controls abschalten. (nur für schnelle Rechner empfohlen)
    SetWinDelay, -1             ; Verzögerung bei allen Fensteroperationen abschalten. (nur für schnelle Rechner empfohlen)

; =='AUTOEXEC-SEKTION'===========================================================================================================

    X := 2015                                              ; Jahreszahl, für die Ostern berechnet werden soll

    K  := Floor(X/100)                                      ; Säkularzahl
    M  := 15+Floor((3*K+3)/4)-Floor((8*K+13)/25)            ; säkulare Mondschaltung
    S  := 2-Floor((3*K+3)/4)                                ; säkulare Sonnenschaltung
    A  := Mod(X,19)                                         ; Mondparameter
    D  := Mod(19*A+M,30)                                    ; Keim für den ersten Vollmond im Frühling
    R  := Floor(D/29)+(Floor(D/28)-Floor(D/29))*Floor(A/11) ; kalendarische Korrekturgröße (beseitigt die Gaußschen Ausnahmeregeln)
    OG := 21+D-R                                            ; Ostergrenze (Märzdatum des Ostervollmonds)
    SZ := 7-Mod(X+Floor(X/4)+S,7)                           ; Erster Sonntag im März
    OE := 7-Mod(OG-SZ,7)                                    ; Entfernung des Ostersonntag von der Ostergrenze in Tagen
    OS := OG+OE                                             ; Märzdatum (ggf. in den April verlängert) des Ostersonntag, (32. März = 1. April usw.)

    FormatTime, Osterdatum,% x (OS > 31 ? "04" SubStr("0"OS-31,-1) : "03" OS), LongDate

    MsgBox, 4160,% "Ostern " X,% Osterdatum