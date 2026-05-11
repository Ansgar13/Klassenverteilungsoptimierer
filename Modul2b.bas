Attribute VB_Name = "Modul2"
Option Explicit
Public ANZAHL_KLASSEN As Integer
Public lastStatus As String
Sub Klassen_Optimierer()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
UserForm1.Show vbModeless
UserForm1.txtStatus.Value = ""



    Dim ws As Worksheet, wsA As Worksheet
    Set ws = Worksheets("alle")
    On Error Resume Next
    Set wsA = Worksheets("andere Empfehlung")
    On Error GoTo 0
    
    '--- Spalten in "alle" finden (Kopf in Zeile 3)
    Dim colName As Long, colVorname As Long
    Dim colM As Long, colW As Long
    Dim colRS As Long, colRSGy As Long, colGy As Long
    Dim colWohnort As Long, colFarbe As Long
    Dim colWunsch As Long, colNicht As Long
    Dim colFix As Long
    Dim colL1 As Long
    
    
Dim wsTiefe As Worksheet


Set wsTiefe = Worksheets("Suchtiefe")

ANZAHL_KLASSEN = wsTiefe.Range("B18").Value

If ANZAHL_KLASSEN < 1 Or ANZAHL_KLASSEN > 7 Then
    MsgBox "Klassenzahl muss zwischen 1 und 7 liegen."
    Exit Sub
End If
    
    
colName = FindeSpalteZ2(ws, "name")
colVorname = FindeSpalteZ2(ws, "vorname")
colM = FindeSpalteZ3(ws, "m")
colW = FindeSpalteZ3(ws, "w")
colRS = FindeSpalteZ3(ws, "rs")
colRSGy = FindeSpalteZ3(ws, "rs/gy")
colGy = FindeSpalteZ3(ws, "gy")
colWohnort = FindeSpalteZ3(ws, "wohnort")
colFarbe = FindeSpalteZ2(ws, "farbe")
colFix = FindeSpalteZ2(ws, "fix")
    Dim colLehrer As Long
colLehrer = FindeSpalte(ws, "klassenlehrer")

If colFix = 0 Then
    MsgBox "Spalte 'Fix' nicht gefunden."
    Exit Sub
End If
Dim colKommentarMit As Long
Dim colKommentarNicht As Long

colKommentarMit = FindeSpalte(ws, "kommentar möchte")
colKommentarNicht = FindeSpalte(ws, "kommentar besser")

If colKommentarMit = 0 Then
    MsgBox "Spalte 'Kommentar möchte mit' nicht gefunden."
    Exit Sub
End If

If colKommentarNicht = 0 Then
    MsgBox "Spalte 'Kommentar besser nicht' nicht gefunden."
    Exit Sub
End If

colWunsch = colKommentarMit + 1
colNicht = FindeSpalte(ws, "besser nicht zusammen mit")


    If colName = 0 Or colVorname = 0 Then
        MsgBox "Spalten 'Name' und/oder 'Vorname' nicht gefunden (Zeile 3).", vbCritical
        Exit Sub
    End If
    
    '--- Klassenlehrer (nur aus Blatt A)
    Dim colLehrerA As Long
    If Not wsA Is Nothing Then
        colLehrerA = FindeSpalteZ2(wsA, "Klassenlehrer")
    End If
    
    '--- Datenbereich
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row
    If lastRow < 4 Then
        MsgBox "Keine Daten ab Zeile 4 gefunden.", vbExclamation
        Exit Sub
    End If
    
    Dim n As Long
    n = lastRow - 3
    
    '--- Arrays
    Dim nameListe() As String
    Dim geschlecht() As String
    Dim schulform() As String
    Dim wohnort() As String
    Dim rot() As Boolean
    Dim wunsch() As String
    Dim nichtMit() As String
    Dim lehrer() As String
    Dim fixed() As Integer
    Dim cluster() As Integer
    
    ReDim nameListe(1 To n)
    ReDim geschlecht(1 To n)
    ReDim schulform(1 To n)
    ReDim wohnort(1 To n)
    ReDim rot(1 To n)
    ReDim wunsch(1 To n)
    ReDim nichtMit(1 To n)
    ReDim lehrer(1 To n)
    ReDim fixed(1 To n)
ReDim cluster(1 To n)
    
    '--- Einlesen aus "alle"
    Dim r As Long, i As Long
    i = 1
    For r = 4 To lastRow
        nameListe(i) = NormalisiereName(ws.Cells(r, colName).Value, ws.Cells(r, colVorname).Value)
        
        If colM > 0 And LCase(Trim(ws.Cells(r, colM).Value)) = "x" Then geschlecht(i) = "m"
        If colW > 0 And LCase(Trim(ws.Cells(r, colW).Value)) = "x" Then geschlecht(i) = "w"
        
        If colRS > 0 And LCase(Trim(ws.Cells(r, colRS).Value)) = "x" Then schulform(i) = "RS"
        If colRSGy > 0 And LCase(Trim(ws.Cells(r, colRSGy).Value)) = "x" Then schulform(i) = "RSGy"
        If colGy > 0 And LCase(Trim(ws.Cells(r, colGy).Value)) = "x" Then schulform(i) = "Gy"
        
        If colWohnort > 0 Then wohnort(i) = Trim(ws.Cells(r, colWohnort).Value)
        
        
        
        If colFarbe > 0 Then rot(i) = (Trim(ws.Cells(r, colFarbe).Value) <> "")
        
        If colLehrer > 0 Then lehrer(i) = Trim(ws.Cells(r, colLehrer).Value)

        
        
        Dim w1 As String
Dim w2 As String

w1 = NormalisiereName(ws.Cells(r, colWunsch).Value, _
                      ws.Cells(r, colWunsch + 1).Value)

w2 = NormalisiereName(ws.Cells(r, colWunsch + 2).Value, _
                      ws.Cells(r, colWunsch + 3).Value)

If w1 = "" And w2 = "" Then
    wunsch(i) = ""
ElseIf w1 = "" Then
    wunsch(i) = w2
ElseIf w2 = "" Then
    wunsch(i) = w1
Else
    wunsch(i) = w1 & "|" & w2
End If
        
        
        Dim n1 As String
Dim n2 As String

n1 = NormalisiereName(ws.Cells(r, colNicht).Value, _
                      ws.Cells(r, colNicht + 1).Value)

n2 = NormalisiereName(ws.Cells(r, colNicht + 2).Value, _
                      ws.Cells(r, colNicht + 3).Value)

If n1 = "" And n2 = "" Then
    nichtMit(i) = ""
ElseIf n1 = "" Then
    nichtMit(i) = n2
ElseIf n2 = "" Then
    nichtMit(i) = n1
Else
    nichtMit(i) = n1 & "|" & n2
End If
        
        
        i = i + 1
    Next r
    
'--- Fixierte Klassen aus Spalte "Fix"

For i = 1 To n

    Dim fixVal As String
    fixVal = LCase(Trim(ws.Cells(i + 3, colFix).Value))

If Left(fixVal, 1) = "5" Then
    fixed(i) = Asc(Right(fixVal, 1)) - 96
End If
Next i

Dim konfliktMap As Object
Set konfliktMap = CreateObject("Scripting.Dictionary")

Dim wunschMap As Object
Set wunschMap = CreateObject("Scripting.Dictionary")

Dim parts() As String
Dim p As Variant

For i = 1 To n

    '--- Konflikte
    If nichtMit(i) <> "" Then
        parts = Split(nichtMit(i), "|")
        konfliktMap(nameListe(i)) = parts
    End If

    '--- Wünsche
    If wunsch(i) <> "" Then
        parts = Split(wunsch(i), "|")
        wunschMap(nameListe(i)) = parts
    End If

Next i





'--- Namensindex für schnelle Suche
Dim nameIndex As Object
Set nameIndex = CreateObject("Scripting.Dictionary")

For i = 1 To n
    nameIndex(nameListe(i)) = i
   
Next i

Call MinMaxCluster_V2(cluster, nichtMit, nameListe, nameIndex)
 
 For i = 1 To n
    Debug.Print nameListe(i), lehrer(i)
Next i

Dim zielwerte As Variant

zielwerte = BerechneZielwerte(geschlecht, schulform, rot, ANZAHL_KLASSEN)

Dim zielM As Double, zielRS As Double, zielRSGy As Double, zielRot As Double

zielM = zielwerte(0)
zielRS = zielwerte(1)
zielRSGy = zielwerte(2)
zielRot = zielwerte(3)



    
'--- Optimierung
Dim klasse() As Integer
ReDim klasse(1 To n)

Dim bestSol() As Integer
ReDim bestSol(1 To 5, 1 To n)

Dim bestScore(1 To 5) As Double
For i = 1 To 5: bestScore(i) = 1E+20: Next i

Randomize

Application.StatusBar = "Optimierung gestartet..."

Dim startTime As Double
startTime = Timer

Dim iter As Long, maxIter As Long
Dim restart As Long, maxRestart As Long

Set wsTiefe = Worksheets("Suchtiefe")

maxIter = wsTiefe.Range("B1").Value
maxRestart = wsTiefe.Range("B3").Value   'NEU: Anzahl Neustarts einstellbar
If maxRestart < 1 Then maxRestart = 1
Dim wKonflikt As Long
Dim wWunsch As Long
Dim wGroesse As Long
Dim wRS As Long
Dim wRSGy As Long
Dim wRot As Long
Dim wGeschlecht As Long
Dim wLehrer As Long

wKonflikt = wsTiefe.Range("B7").Value
wWunsch = wsTiefe.Range("B8").Value
wGroesse = wsTiefe.Range("B9").Value
wRS = wsTiefe.Range("B10").Value
wRSGy = wsTiefe.Range("B11").Value
wRot = wsTiefe.Range("B12").Value
wGeschlecht = wsTiefe.Range("B13").Value
wLehrer = wsTiefe.Range("B14").Value

'====================================================
' Random Restarts
'====================================================



For restart = 1 To maxRestart
UserForm1.AddStatus "=== Restart " & restart & " ==="
DoEvents
    If restart = 1 Then
    
        'Erster Start: gute Lösung
        Call GreedyStart(klasse, nameListe, geschlecht, schulform, rot, wohnort, _
                     wunsch, nichtMit, lehrer, nameIndex, _
                     wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                     fixed, cluster, ANZAHL_KLASSEN, _
                     zielM, zielRS, zielRSGy, zielRot)
    
    ElseIf Rnd < 0.7 Then
    
        '70%: komplett neu starten
        If Rnd < 0.5 Then
       Call GreedyStart(klasse, nameListe, geschlecht, schulform, rot, wohnort, _
                     wunsch, nichtMit, lehrer, nameIndex, _
                     wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                     fixed, cluster, ANZAHL_KLASSEN, _
                     zielM, zielRS, zielRSGy, zielRot)
   
            
        Else
                Call RandomStart(klasse, fixed, cluster, ANZAHL_KLASSEN)

        End If
    
    Else
    
        '30%: BESTE Lösung weiter verbessern ??
        For i = 1 To n
            klasse(i) = bestSol(1, i)
        Next i
    
    End If









For iter = 1 To maxIter

        If iter Mod 50 = 0 Then
        
lastStatus = "Restart " & restart & "/" & maxRestart & _
             "   Iteration " & iter & "/" & maxIter & _
             "   Zeit: " & Format(Timer - startTime, "0.0") & "s"
             
           If iter Mod 50 = 0 Then
    UserForm1.txtStatus.Value = lastStatus
    DoEvents
End If
           
           
        End If

        'lokale Verbesserung
Call LocalImprove(klasse, nameListe, geschlecht, schulform, rot, wohnort, _
                  wunsch, nichtMit, lehrer, _
                  nameIndex, konfliktMap, wunschMap, _
                  wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, wLehrer, _
                  fixed, cluster, ANZAHL_KLASSEN, _
                  zielM, zielRS, zielRSGy, zielRot)
        
        
        
        Dim sc As Double




sc = BewertungFast(klasse, nameListe, geschlecht, schulform, rot, _
                   nameIndex, konfliktMap, wunschMap, _
                   wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                   ANZAHL_KLASSEN, _
                   zielM, zielRS, zielRSGy, zielRot)
                   
                   Dim k As Integer

        For k = 1 To 5
            If sc < bestScore(k) Then
                bestScore(k) = sc
                For i = 1 To n
                    bestSol(k, i) = klasse(i)
                Next i
                Exit For
            End If
        Next k

    Next iter
Next restart
    
Dim s2 As Integer

'--- Original-Fix sichern
Dim fixedOriginal() As Integer
ReDim fixedOriginal(1 To n)

For i = 1 To n
    fixedOriginal(i) = fixed(i)
Next i

'========================================
' PHASE 2: Fix & Repair für ALLE Lösungen
'========================================

For s2 = 1 To 5

    '--- Fix zurücksetzen
    For i = 1 To n
        fixed(i) = fixedOriginal(i)
    Next i

    Dim problem() As Boolean
    ReDim problem(1 To n)

    '--- Lösung s2 laden
    For i = 1 To n
        klasse(i) = bestSol(s2, i)
    Next i
    
    
    
'--- Wunschverletzungen markieren

Dim j As Long
Dim erfuellt As Boolean

For i = 1 To n

    If wunsch(i) <> "" Then
    
        parts = Split(wunsch(i), "|")
        erfuellt = False
        
        For Each p In parts
        
            If Trim(p) <> "" Then
            
                If nameIndex.Exists(p) Then
                
                    j = nameIndex(p)
                    
                    If klasse(i) = klasse(j) Then
                        erfuellt = True
                        Exit For
                    End If
                    
                End If
                
            End If
            
        Next p
        
        If Not erfuellt Then
            problem(i) = True
        End If
        
    End If

Next i
    
 '--- Konflikte markieren

Dim parts2() As String
Dim p2 As Variant
Dim j2 As Long

For i = 1 To n

    If nichtMit(i) <> "" Then
    
        parts2 = Split(nichtMit(i), "|")
        
        For Each p2 In parts2
        
            If Trim(p2) <> "" Then
            
                If nameIndex.Exists(p2) Then
                
                    j2 = nameIndex(p2)
                    
                    If klasse(i) = klasse(j2) Then
                        problem(i) = True
                        problem(j2) = True   '?? ganz wichtig!
                    End If
                    
                End If
                
            End If
            
        Next p2
        
    End If

Next i
    
'--- Fix setzen (nur Problem-Schüler beweglich)

Dim clusterSize() As Long
ReDim clusterSize(1 To n)

Dim ci As Long

For i = 1 To n
    ci = cluster(i)
    If ci > 0 Then
        clusterSize(ci) = clusterSize(ci) + 1
    End If
Next i



For i = 1 To n

    If problem(i) = False Then
        fixed(i) = klasse(i)
    Else
        fixed(i) = 0
    End If

Next i



   '========================================
' PHASE 2: gezielte Nachoptimierung
'========================================

Dim iter2 As Long
Dim maxIter2 As Long

maxIter2 = wsTiefe.Range("B1").Value / 2   'z.B. halbe Iterationen

'Gewichte optional verstärken
Dim wWunsch2 As Long, wKonflikt2 As Long
wWunsch2 = wWunsch * 2
wKonflikt2 = wKonflikt * 2


For iter2 = 1 To maxIter2

    If iter2 Mod 50 = 0 Then
        UserForm1.AddStatus "Phase 2 - Iteration " & iter2 & "/" & maxIter2
        DoEvents
    End If
Call LocalImprove(klasse, nameListe, geschlecht, schulform, rot, wohnort, _
                  wunsch, nichtMit, lehrer, _
                  nameIndex, konfliktMap, wunschMap, _
                  wKonflikt2, wWunsch2, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                  wLehrer, fixed, cluster, ANZAHL_KLASSEN, _
                  zielM, zielRS, zielRSGy, zielRot)
Next iter2
    
    
        '--- Lösung zurückspeichern
    For i = 1 To n
        bestSol(s2, i) = klasse(i)
    Next i

Next s2
    
    
    
    
    
    
    
    Dim klasseName() As String
ReDim klasseName(1 To ANZAHL_KLASSEN)

Dim iK As Integer
For iK = 1 To ANZAHL_KLASSEN
    klasseName(iK) = "5" & Chr(96 + iK)   'a,b,c,d,e,f,g
Next iK
    
    '--- L1..L5 Spalten finden/erzeugen
    
    colL1 = FindeSpalte(ws, "l1")
    If colL1 = 0 Then
        colL1 = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(3, colL1).Value = "L1"
        ws.Cells(3, colL1 + 1).Value = "L2"
        ws.Cells(3, colL1 + 2).Value = "L3"
        ws.Cells(3, colL1 + 3).Value = "L4"
        ws.Cells(3, colL1 + 4).Value = "L5"
    End If
    
    '--- Klassenbezeichnungen
    
    
    
    '--- Schreiben in L1..L5 (überschreibt)
   Dim rowIndex As Long
Dim s As Integer
    For s = 1 To 5
        rowIndex = 4
        For i = 1 To n
            ws.Cells(rowIndex, colL1 + s - 1).Value = klasseName(bestSol(s, i))
            rowIndex = rowIndex + 1
        Next i
    Next s
Call AnalyseErstellen(ws, nameListe, geschlecht, schulform, rot, wohnort, wunsch, nichtMit, lehrer, colL1, n, ANZAHL_KLASSEN)
Call WunschFehlerListe(ws, nameListe, wunsch, colL1, n)
Call KonfliktListe(ws, nameListe, nichtMit, colL1, n)
    
    Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
    
    MsgBox "Optimierung abgeschlossen. Lösungen in L1–L5.", vbInformation
Application.StatusBar = False
End Sub

Sub RandomStart(ByRef klasse() As Integer, _
                ByRef fixed() As Integer, _
                ByRef cluster() As Integer, _
                ByVal ANZAHL_KLASSEN As Integer)

    Dim i As Long
    Dim cl As Long
    Dim c As Integer
    
    'Max Cluster bestimmen
    Dim maxCluster As Long
    For i = 1 To UBound(cluster)
        If cluster(i) > maxCluster Then maxCluster = cluster(i)
    Next i

    'Clusterweise zufällig zuweisen
    For cl = 1 To maxCluster

        'zufällige Klasse
        c = 1 + Int(Rnd * ANZAHL_KLASSEN)

        For i = 1 To UBound(cluster)

            If cluster(i) = cl Then

                'Fixierte behalten ihre Klasse
                If fixed(i) > 0 Then
                    klasse(i) = fixed(i)
                Else
                    klasse(i) = c
                End If

            End If

        Next i

    Next cl

End Sub


Sub GreedyStart(ByRef klasse() As Integer, _
                ByRef nameListe() As String, _
                ByRef geschlecht() As String, _
                ByRef schulform() As String, _
                ByRef rot() As Boolean, _
                ByRef wohnort() As String, _
                ByRef wunsch() As String, _
                ByRef nichtMit() As String, _
                ByRef lehrer() As String, _
                ByVal nameIndex As Object, _
                ByVal wKonflikt As Long, _
                ByVal wWunsch As Long, _
                ByVal wGroesse As Long, _
                ByVal wRS As Long, _
                ByVal wRSGy As Long, _
                ByVal wRot As Long, _
                ByVal wGeschlecht As Long, _
                ByRef fixed() As Integer, _
                ByRef cluster() As Integer, _
                ByVal ANZAHL_KLASSEN As Integer, _
                ByVal zielM As Double, _
                ByVal zielRS As Double, _
                ByVal zielRSGy As Double, _
                ByVal zielRot As Double)

Dim i As Long, c As Integer, cl As Long
Dim bestClass As Integer
Dim bestScore As Double
Dim sc As Double
Dim maxCluster As Long

Dim size() As Long
Dim m() As Long
Dim rs() As Long
Dim rsgy() As Long
Dim rotc() As Long

ReDim size(1 To ANZAHL_KLASSEN)
ReDim m(1 To ANZAHL_KLASSEN)
ReDim rs(1 To ANZAHL_KLASSEN)
ReDim rsgy(1 To ANZAHL_KLASSEN)
ReDim rotc(1 To ANZAHL_KLASSEN)

'------------------------------------------
' Max Cluster bestimmen
'------------------------------------------
For i = 1 To UBound(cluster)
    If cluster(i) > maxCluster Then maxCluster = cluster(i)
Next i

'------------------------------------------
' Fixierte Schüler setzen
'------------------------------------------
For i = 1 To UBound(klasse)

    If fixed(i) > 0 Then
    
        klasse(i) = fixed(i)
        
        size(fixed(i)) = size(fixed(i)) + 1
        
        If geschlecht(i) = "m" Then m(fixed(i)) = m(fixed(i)) + 1
        If schulform(i) = "RS" Then rs(fixed(i)) = rs(fixed(i)) + 1
        If schulform(i) = "RSGy" Then rsgy(fixed(i)) = rsgy(fixed(i)) + 1
        If rot(i) Then rotc(fixed(i)) = rotc(fixed(i)) + 1
        
    End If

Next i

'------------------------------------------
' Clusterliste mischen
'------------------------------------------
Dim order() As Long
ReDim order(1 To maxCluster)

For cl = 1 To maxCluster
    order(cl) = cl
Next cl

Dim idx As Long, r As Long, tmp As Long

For idx = 1 To maxCluster
    r = idx + Int(Rnd * (maxCluster - idx + 1))
    tmp = order(idx)
    order(idx) = order(r)
    order(r) = tmp
Next idx

'------------------------------------------
' Cluster verteilen
'------------------------------------------
For idx = 1 To maxCluster

    cl = order(idx)

    Dim clusterSize As Long
    Dim clusterM As Long
    Dim clusterRS As Long
    Dim clusterRSGy As Long
    Dim clusterRot As Long
    
    clusterSize = 0
    clusterM = 0
    clusterRS = 0
    clusterRSGy = 0
    clusterRot = 0

    For i = 1 To UBound(cluster)
    
        If cluster(i) = cl Then
        
            clusterSize = clusterSize + 1
            
            If geschlecht(i) = "m" Then clusterM = clusterM + 1
            If schulform(i) = "RS" Then clusterRS = clusterRS + 1
            If schulform(i) = "RSGy" Then clusterRSGy = clusterRSGy + 1
            If rot(i) Then clusterRot = clusterRot + 1
            
        End If
        
    Next i

    '------------------------------------------
    ' Fixklasse prüfen
    '------------------------------------------
    Dim fixedClass As Integer
    fixedClass = 0

    For i = 1 To UBound(cluster)

        If cluster(i) = cl And fixed(i) > 0 Then
        
            fixedClass = fixed(i)
            Exit For
            
        End If

    Next i

    '------------------------------------------
    ' Zielklasse bestimmen
    '------------------------------------------
    If fixedClass > 0 Then
    
        bestClass = fixedClass
        
    Else
    
        bestScore = 1E+20
        
        For c = 1 To ANZAHL_KLASSEN
        
            Dim ziel As Double
            ziel = UBound(klasse) / ANZAHL_KLASSEN
            
            sc = 0
            
            sc = sc + Abs((size(c) + clusterSize) - ziel) * wGroesse
sc = sc + Abs((m(c) + clusterM) - zielM) * wGeschlecht
sc = sc + Abs((rs(c) + clusterRS) - zielRS) * wRS
sc = sc + Abs((rsgy(c) + clusterRSGy) - zielRSGy) * wRSGy
sc = sc + Abs((rotc(c) + clusterRot) - zielRot) * wRot
            If sc < bestScore Then
                bestScore = sc
                bestClass = c
            End If
            
        Next c
        
        If bestClass = 0 Then bestClass = 1
        
    End If

    '------------------------------------------
    ' Cluster setzen
    '------------------------------------------
    For i = 1 To UBound(cluster)

        If cluster(i) = cl Then
        
            If fixed(i) = 0 Then
                klasse(i) = bestClass
            End If
            
            size(bestClass) = size(bestClass) + 1
            
            If geschlecht(i) = "m" Then m(bestClass) = m(bestClass) + 1
            If schulform(i) = "RS" Then rs(bestClass) = rs(bestClass) + 1
            If schulform(i) = "RSGy" Then rsgy(bestClass) = rsgy(bestClass) + 1
            If rot(i) Then rotc(bestClass) = rotc(bestClass) + 1
            
        End If

    Next i

Next idx

End Sub
'--- Lokale Verbesserungen durch Tauschen
Sub LocalImprove(ByRef k() As Integer, ByRef nameListe() As String, ByRef geschlecht() As String, _
                 ByRef schulform() As String, ByRef rot() As Boolean, ByRef wohnort() As String, _
                 ByRef wunsch() As String, ByRef nichtMit() As String, ByRef lehrer() As String, _
                 ByVal nameIndex As Object, ByVal konfliktMap As Object, ByVal wunschMap As Object, _
                 ByVal wKonflikt As Long, ByVal wWunsch As Long, ByVal wGroesse As Long, ByVal wRS As Long, _
                 ByVal wRSGy As Long, ByVal wRot As Long, ByVal wGeschlecht As Long, ByVal wLehrer As Long, _
                 ByRef fixed() As Integer, ByRef cluster() As Integer, ByVal ANZAHL_KLASSEN As Integer, _
                 ByVal zielM As Double, ByVal zielRS As Double, ByVal zielRSGy As Double, ByVal zielRot As Double)

                 
        
                 
Dim wsTiefe As Worksheet
Set wsTiefe = Worksheets("Suchtiefe")

Dim tauschVersuche As Long
tauschVersuche = wsTiefe.Range("B2").Value

Dim i As Long
Dim a As Long, b As Long
Dim t As Integer

Dim curScore As Double
Dim newScore As Double

Dim temperature As Double
temperature = 3000

curScore = BewertungFast(k, nameListe, geschlecht, schulform, rot, _
                         nameIndex, konfliktMap, wunschMap, _
                         wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                         ANZAHL_KLASSEN, _
                         zielM, zielRS, zielRSGy, zielRot)
                         
                         
 If wGroesse > 0 Then
    Dim toleranz As Long
    If IsEmpty(wsTiefe.Range("B19").Value) Or wsTiefe.Range("B19").Value = "" Then
        toleranz = 3   'Standardwert falls B19 leer
    Else
        toleranz = CLng(wsTiefe.Range("B19").Value)
    End If
    If toleranz < 0 Then toleranz = 0
    Call KlassenGroesseKorrigieren(k, cluster, fixed, ANZAHL_KLASSEN, toleranz)
End If

For i = 1 To tauschVersuche

'------------------------------------------
' NEU: Einzel-Schüler verschieben (wichtig!)
'------------------------------------------
Dim moveIndex As Long
moveIndex = 1 + Int(Rnd * UBound(k))

If fixed(moveIndex) = 0 Then

    Dim oldClass As Integer
    Dim newClass As Integer
    
    oldClass = k(moveIndex)
    newClass = 1 + Int(Rnd * ANZAHL_KLASSEN)
    
    If newClass <> oldClass Then
    
        k(moveIndex) = newClass
        
        newScore = BewertungFast(k, nameListe, geschlecht, schulform, rot, _
                         nameIndex, konfliktMap, wunschMap, _
                         wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                         ANZAHL_KLASSEN, _
                         zielM, zielRS, zielRSGy, zielRot)
        
        If newScore < curScore Then
            curScore = newScore
        Else
            k(moveIndex) = oldClass ' Undo
        End If
        
    End If
    
End If


a = 1 + Int(Rnd * UBound(k))
b = 1 + Int(Rnd * UBound(k))

If k(a) = k(b) Then GoTo NextTry

If a = b Then GoTo NextTry

Dim ca As Long
Dim cb As Long

ca = cluster(a)
cb = cluster(b)

If ca = cb Then GoTo NextTry

Dim i2 As Long
Dim classA As Integer
Dim classB As Integer

classA = k(a)
classB = k(b)

For i2 = 1 To UBound(k)

    If cluster(i2) = ca Then
        If fixed(i2) > 0 Then GoTo NextTry
    End If
    
    If cluster(i2) = cb Then
        If fixed(i2) > 0 Then GoTo NextTry
    End If

Next i2

For i2 = 1 To UBound(k)

    If cluster(i2) = ca Then k(i2) = classB
    If cluster(i2) = cb Then k(i2) = classA

Next i2


newScore = BewertungFast(k, nameListe, geschlecht, schulform, rot, _
                         nameIndex, konfliktMap, wunschMap, _
                         wKonflikt, wWunsch, wGroesse, wRS, wRSGy, wRot, wGeschlecht, _
                         ANZAHL_KLASSEN, _
                         zielM, zielRS, zielRSGy, zielRot)
                     
                     
                     
    If newScore < curScore Then

        curScore = newScore

    Else

        Dim delta As Double
        delta = newScore - curScore

        If Rnd < Exp(-delta / temperature) Then
            curScore = newScore
        Else
            'Undo
          For i2 = 1 To UBound(k)

    If cluster(i2) = ca Then k(i2) = classA
    If cluster(i2) = cb Then k(i2) = classB

Next i2
        End If

    End If

    temperature = temperature * 0.995

NextTry:
Next i

End Sub
'--- Bewertungsfunktion
Function BewertungFast(ByRef k() As Integer, _
                       ByRef nameListe() As String, _
                       ByRef geschlecht() As String, _
                       ByRef schulform() As String, _
                       ByRef rot() As Boolean, _
                       ByVal nameIndex As Object, _
                       ByVal konfliktMap As Object, _
                       ByVal wunschMap As Object, _
                       ByVal wKonflikt As Long, _
                       ByVal wWunsch As Long, _
                       ByVal wGroesse As Long, _
                       ByVal wRS As Long, _
                       ByVal wRSGy As Long, _
                       ByVal wRot As Long, _
                       ByVal wGeschlecht As Long, _
                       ByVal ANZAHL_KLASSEN As Integer, _
                       ByVal zielM As Double, _
                       ByVal zielRS As Double, _
                       ByVal zielRSGy As Double, _
                       ByVal zielRot As Double) As Double

Dim score As Double
Dim size() As Long, m() As Long, rs() As Long, rsgy() As Long, rotc() As Long
Dim i As Long, c As Integer

ReDim size(1 To ANZAHL_KLASSEN)
ReDim m(1 To ANZAHL_KLASSEN)
ReDim rs(1 To ANZAHL_KLASSEN)
ReDim rsgy(1 To ANZAHL_KLASSEN)
ReDim rotc(1 To ANZAHL_KLASSEN)

'----------------------------------
' Statistik sammeln
'----------------------------------
For i = 1 To UBound(k)

    c = k(i)
    If c < 1 Or c > ANZAHL_KLASSEN Then GoTo NextI

    size(c) = size(c) + 1
    
    If geschlecht(i) = "m" Then m(c) = m(c) + 1
    If schulform(i) = "RS" Then rs(c) = rs(c) + 1
    If schulform(i) = "RSGy" Then rsgy(c) = rsgy(c) + 1
    If rot(i) Then rotc(c) = rotc(c) + 1

NextI:
Next i

'----------------------------------
' Klassenbewertung
'----------------------------------
Dim ziel As Double
ziel = UBound(k) / ANZAHL_KLASSEN

For c = 1 To ANZAHL_KLASSEN

    score = score + Abs(size(c) - ziel) * wGroesse
    score = score + Abs(m(c) - zielM) * wGeschlecht
    score = score + Abs(rs(c) - zielRS) * wRS
    score = score + Abs(rsgy(c) - zielRSGy) * wRSGy
    score = score + Abs(rotc(c) - zielRot) * wRot

Next c

'----------------------------------
' Wünsche (SCHNELL!)
'----------------------------------
Dim p As Variant, j As Long
Dim erfüllt As Boolean

For i = 1 To UBound(k)

    If wunschMap.Exists(nameListe(i)) Then
    
        erfüllt = False
        
        For Each p In wunschMap(nameListe(i))
        
            If nameIndex.Exists(p) Then
            
                j = nameIndex(p)
                
                If k(i) = k(j) Then
                    erfüllt = True
                    Exit For
                End If
                
            End If
            
        Next p
        
        If Not erfüllt Then score = score + wWunsch
        
    End If

Next i

'----------------------------------
' Konflikte (SCHNELL!)
'----------------------------------
For i = 1 To UBound(k)

    If konfliktMap.Exists(nameListe(i)) Then
    
        For Each p In konfliktMap(nameListe(i))
        
            If nameIndex.Exists(p) Then
            
                j = nameIndex(p)
                
                If k(i) = k(j) Then
                    score = score + wKonflikt
                End If
                
            End If
            
        Next p
        
    End If

Next i

BewertungFast = score

End Function

'--- Spalte finden (Zeile 3)
Function FindeSpalte(ws As Worksheet, suchText As String, Optional startCol As Long = 1) As Long

Dim c As Long, r As Long
Dim lastCol As Long

lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column

For c = startCol To lastCol
    For r = 1 To 3
    
        If InStr(LCase(ws.Cells(r, c).Value), LCase(suchText)) > 0 Then
            FindeSpalte = c
            Exit Function
        End If
        
    Next r
Next c

FindeSpalte = 0

End Function
Function FindeSpalteZ2(ws As Worksheet, suchText As String) As Long

Dim c As Long
Dim lastCol As Long
Dim val As String

lastCol = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

For c = 1 To lastCol

    val = LCase(Trim(ws.Cells(2, c).Value))
    
    If val = LCase(suchText) Then
        FindeSpalteZ2 = c
        Exit Function
    End If

Next c

FindeSpalteZ2 = 0

End Function
Function FindeSpalteZ3(ws As Worksheet, suchText As String) As Long

Dim c As Long
Dim lastCol As Long
Dim val As String

lastCol = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Column

For c = 1 To lastCol

    val = LCase(Trim(ws.Cells(3, c).Value))
    
    If val = LCase(suchText) Then
        FindeSpalteZ3 = c
        Exit Function
    End If

Next c

FindeSpalteZ3 = 0

End Function
'--- Namen normalisieren
Function NormalisiereName(nachname As String, vorname As String) As String
    Dim t As String
    t = Trim(nachname & " " & vorname)
    t = Replace(t, ",", " ")
    t = Application.WorksheetFunction.Trim(t)
    NormalisiereName = t
End Function


'--- Fuzzy-Namensvergleich
Function NamenMatch(name1 As String, name2 As String) As Boolean
    Dim n1 As String, n2 As String
    n1 = LCase(Replace(name1, ",", ""))
    n2 = LCase(Replace(name2, ",", ""))
    If InStr(n1, n2) > 0 Or InStr(n2, n1) > 0 Then
        NamenMatch = True
    Else
        NamenMatch = False
    End If
End Function


Sub AnalyseErstellen(ws As Worksheet, _
                     nameListe() As String, _
                     geschlecht() As String, _
                     schulform() As String, _
                     rot() As Boolean, _
                     wohnort() As String, _
                     wunsch() As String, _
                     nichtMit() As String, _
                     lehrer() As String, _
                     colL1 As Long, _
                     n As Long, _
                     ANZAHL_KLASSEN As Integer)



Dim wsA As Worksheet

On Error Resume Next
Set wsA = Worksheets("Analyse")
On Error GoTo 0

If wsA Is Nothing Then
    Set wsA = Worksheets.Add
    wsA.Name = "Analyse"
Else
    wsA.Cells.Clear
End If

Dim klasseName() As String
ReDim klasseName(1 To ANZAHL_KLASSEN)

Dim iK As Integer
For iK = 1 To ANZAHL_KLASSEN
    klasseName(iK) = "5" & Chr(96 + iK)   'a,b,c,d,e,f,g
Next iK

wsA.Cells(1, 1) = "Lösung"
wsA.Cells(1, 2) = "Klasse"
wsA.Cells(1, 3) = "Größe"
wsA.Cells(1, 4) = "m"
wsA.Cells(1, 5) = "w"
wsA.Cells(1, 6) = "RS"
wsA.Cells(1, 7) = "RS/Gy"
wsA.Cells(1, 8) = "Rot"
wsA.Cells(1, 9) = "Wunschfehler"
wsA.Cells(1, 10) = "Konflikte"
wsA.Cells(1, 11) = "Max Lehrergruppe"
Dim rowOut As Long
rowOut = 2

Dim s As Integer, c As Integer, i As Long
Dim klasse As String

For s = 1 To 5
Dim wunschFehler As Long
Dim konfliktFehler As Long

wunschFehler = 0
konfliktFehler = 0
    For c = 1 To ANZAHL_KLASSEN

        Dim size As Long
        Dim m As Long
        Dim w As Long
        Dim rs As Long
        Dim rsgy As Long
        Dim rotc As Long
        Dim teacherCount As Object
Set teacherCount = CreateObject("Scripting.Dictionary")
        
        'WICHTIG: Werte zurücksetzen
        size = 0
        m = 0
        w = 0
        rs = 0
        rsgy = 0
        rotc = 0

        For i = 1 To n

            klasse = ws.Cells(i + 3, colL1 + s - 1).Value

            If klasse = klasseName(c) Then
                If lehrer(i) <> "" Then

    If teacherCount.Exists(lehrer(i)) Then
        teacherCount(lehrer(i)) = teacherCount(lehrer(i)) + 1
    Else
        teacherCount(lehrer(i)) = 1
    End If

End If
                size = size + 1

                If geschlecht(i) = "m" Then m = m + 1
                If geschlecht(i) = "w" Then w = w + 1

                If schulform(i) = "RS" Then rs = rs + 1
                If schulform(i) = "RSGy" Then rsgy = rsgy + 1

                If rot(i) Then rotc = rotc + 1

            End If

        Next i
        Dim maxGroup As Long
Dim t As Variant

maxGroup = 0

For Each t In teacherCount.keys

    If teacherCount(t) > maxGroup Then
        maxGroup = teacherCount(t)
    End If

Next t
        wsA.Cells(rowOut, 1) = "L" & s
        wsA.Cells(rowOut, 2) = klasseName(c)
        wsA.Cells(rowOut, 3) = size
        wsA.Cells(rowOut, 4) = m
        wsA.Cells(rowOut, 5) = w
        wsA.Cells(rowOut, 6) = rs
        wsA.Cells(rowOut, 7) = rsgy
        wsA.Cells(rowOut, 8) = rotc
        wsA.Cells(rowOut, 11) = maxGroup
        
        rowOut = rowOut + 1

    Next c


'Wunschfehler zählen
Dim i2 As Long, j2 As Long
Dim partnerGefunden As Boolean

For i2 = 1 To n

    If Trim(wunsch(i2)) <> "" Then

        partnerGefunden = False

Dim parts() As String
Dim p As Variant

parts = Split(wunsch(i2), "|")

For Each p In parts

    If Trim(p) <> "" Then

        For j2 = 1 To n

            If NamenMatch(CStr(p), nameListe(j2)) Then

                If ws.Cells(i2 + 3, colL1 + s - 1) = ws.Cells(j2 + 3, colL1 + s - 1) Then
                    partnerGefunden = True
                    Exit For
                End If

            End If

        Next j2

    End If

Next p
        If partnerGefunden = False Then
            wunschFehler = wunschFehler + 1
        End If

    End If

Next i2


Dim parts2() As String
Dim p2 As Variant

For i2 = 1 To n - 1

    If nichtMit(i2) <> "" Then
    
        parts2 = Split(nichtMit(i2), "|")
        
        For Each p2 In parts2
        
            For j2 = i2 + 1 To n
            
                If NamenMatch(CStr(p2), nameListe(j2)) Then
                
                    If ws.Cells(i2 + 3, colL1 + s - 1) = ws.Cells(j2 + 3, colL1 + s - 1) Then
                        konfliktFehler = konfliktFehler + 1
                    End If
                    
                End If
                
            Next j2
            
        Next p2
        
    End If

Next i2

'Ergebnis der Lösung in Analyse schreiben
wsA.Cells(rowOut - ANZAHL_KLASSEN, 9) = wunschFehler
wsA.Cells(rowOut - ANZAHL_KLASSEN, 10) = konfliktFehler

rowOut = rowOut + 1   ' <-- NEU: Leerzeile zwischen Lösungen

Next s
wsA.Columns.AutoFit

End Sub

Sub WunschFehlerListe(ws As Worksheet, _
                      ByRef nameListe() As String, _
                      ByRef wunsch() As String, _
                      colL1 As Long, _
                      n As Long)
Dim wsF As Worksheet

On Error Resume Next
Set wsF = Worksheets("Wunschfehler")
On Error GoTo 0

If wsF Is Nothing Then
    Set wsF = Worksheets.Add
    wsF.Name = "Wunschfehler"
Else
    wsF.Cells.Clear
End If

wsF.Cells(1, 1) = "Lösung"
wsF.Cells(1, 2) = "Zeile in 'alle'"
wsF.Cells(1, 3) = "Name"
wsF.Cells(1, 4) = "Wunsch"

Dim rowOut As Long
rowOut = 2

Dim s As Integer
Dim i As Long, j As Long
Dim partnerGefunden As Boolean

For s = 1 To 5

    For i = 1 To n

        If Trim(wunsch(i)) <> "" Then
        
            partnerGefunden = False
        
            For j = 1 To n
            
                If NamenMatch(wunsch(i), nameListe(j)) Then
                
                    If ws.Cells(i + 3, colL1 + s - 1) = ws.Cells(j + 3, colL1 + s - 1) Then
                        partnerGefunden = True
                        Exit For
                    End If
                    
                End If
                
            Next j
            
            If partnerGefunden = False Then
            
                wsF.Cells(rowOut, 1) = "L" & s
                wsF.Cells(rowOut, 2) = i + 3
                wsF.Cells(rowOut, 3) = nameListe(i)
                wsF.Cells(rowOut, 4) = wunsch(i)
                
                rowOut = rowOut + 1
                
            End If
            
        End If

    Next i
rowOut = rowOut + 1
Next s

wsF.Columns.AutoFit

End Sub

Sub KonfliktListe(ws As Worksheet, _
                  ByRef nameListe() As String, _
                  ByRef nichtMit() As String, _
                  colL1 As Long, _
                  n As Long)

Dim wsK As Worksheet

On Error Resume Next
Set wsK = Worksheets("Konflikte")
On Error GoTo 0

If wsK Is Nothing Then
    Set wsK = Worksheets.Add
    wsK.Name = "Konflikte"
Else
    wsK.Cells.Clear
End If

wsK.Cells(1, 1) = "Lösung"
wsK.Cells(1, 2) = "Zeile 1"
wsK.Cells(1, 3) = "Name 1"
wsK.Cells(1, 4) = "Zeile 2"
wsK.Cells(1, 5) = "Name 2"

Dim rowOut As Long
rowOut = 2

Dim s As Integer
Dim i As Long, j As Long
Dim parts2() As String
Dim p2 As Variant

For s = 1 To 5

    For i = 1 To n

        If nichtMit(i) <> "" Then

            parts2 = Split(nichtMit(i), "|")

            For Each p2 In parts2

                For j = 1 To n

                    If NamenMatch(CStr(p2), nameListe(j)) Then

                        If ws.Cells(i + 3, colL1 + s - 1) = ws.Cells(j + 3, colL1 + s - 1) Then

                            wsK.Cells(rowOut, 1) = "L" & s
                            wsK.Cells(rowOut, 2) = i + 3
                            wsK.Cells(rowOut, 3) = nameListe(i)
                            wsK.Cells(rowOut, 4) = j + 3
                            wsK.Cells(rowOut, 5) = nameListe(j)

                            rowOut = rowOut + 1

                        End If

                    End If

                Next j

            Next p2

        End If

    Next i
rowOut = rowOut + 1   'Leerzeile zwischen Lösungen

Next s

wsK.Columns.AutoFit

End Sub

Sub KlassenGroesseKorrigieren(k() As Integer, cluster() As Integer, fixed() As Integer, ANZAHL_KLASSEN As Integer, toleranz As Long)

    Dim size() As Long
    ReDim size(1 To ANZAHL_KLASSEN)

    Dim i As Long, c As Integer

    '--- Klassengrößen zählen
    For i = 1 To UBound(k)
        size(k(i)) = size(k(i)) + 1
    Next i

    Dim ziel As Double
    ziel = UBound(k) / ANZAHL_KLASSEN

    Dim cBig As Integer, cSmall As Integer
    Dim moved As Boolean
    
    'Mehrfach versuchen, auszugleichen
    Dim iter As Long
    For iter = 1 To 50

        moved = False

        '--- größte Klasse suchen
        cBig = 1
        For c = 2 To ANZAHL_KLASSEN
            If size(c) > size(cBig) Then cBig = c
        Next c

        '--- kleinste Klasse suchen
        cSmall = 1
        For c = 2 To ANZAHL_KLASSEN
            If size(c) < size(cSmall) Then cSmall = c
        Next c

        '--- Abbruch, wenn schon gut verteilt (Toleranz aus Suchtiefe B19)
        If size(cBig) <= ziel + toleranz And size(cSmall) >= ziel - toleranz Then Exit For

        '--- Schüler von groß nach klein verschieben
        For i = 1 To UBound(k)

            If k(i) = cBig And fixed(i) = 0 Then

                k(i) = cSmall

                size(cBig) = size(cBig) - 1
                size(cSmall) = size(cSmall) + 1

                moved = True
                Exit For

            End If

        Next i

        '--- Wenn nichts mehr verschiebbar -> abbrechen
        If Not moved Then Exit For

    Next iter

End Sub

Sub UngueltigeNamenRot()

Dim ws As Worksheet
Set ws = Worksheets("alle")

Dim colName As Long, colVorname As Long
Dim colWunsch As Long, colNicht As Long
Dim colKommentarMit As Long, colKommentarNicht As Long

Dim lastRow As Long
Dim r As Long, r2 As Long, c As Long

Dim nach As String, vor As String
Dim gefunden As Boolean

'Spalten der Schülerliste finden
colName = FindeSpalteZ2(ws, "name")
colVorname = FindeSpalteZ2(ws, "vorname")

'Spalten der Wunschfelder finden
colKommentarMit = FindeSpalte(ws, "kommentar möchte")
colKommentarNicht = FindeSpalte(ws, "kommentar besser")

colWunsch = colKommentarMit + 1
colNicht = colKommentarNicht + 1

If colWunsch = 0 And colNicht = 0 Then
    MsgBox "Keine der gewünschten Spalten gefunden."
    Exit Sub
End If

'letzte Schülerzeile
lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row

For r = 4 To lastRow

'================================================
' möchte gerne zusammen mit
'================================================
If colWunsch > 0 Then

    For c = 0 To 1

        nach = Trim(ws.Cells(r, colWunsch + c * 2).Value)
        vor = Trim(ws.Cells(r, colWunsch + c * 2 + 1).Value)

        If nach <> "" And vor <> "" Then

            gefunden = False

            For r2 = 4 To lastRow

                If Trim(ws.Cells(r2, colName).Value) = nach _
                And Trim(ws.Cells(r2, colVorname).Value) = vor Then

                    gefunden = True
                    Exit For

                End If

            Next r2

            If Not gefunden Then
                ws.Cells(r, colWunsch + c * 2).Interior.Color = RGB(255, 0, 0)
                ws.Cells(r, colWunsch + c * 2 + 1).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(r, colWunsch + c * 2).Interior.ColorIndex = xlNone
                ws.Cells(r, colWunsch + c * 2 + 1).Interior.ColorIndex = xlNone
            End If

        End If

    Next c

End If

'================================================
' besser nicht zusammen mit
'================================================
If colNicht > 0 Then

    For c = 0 To 1

        nach = Trim(ws.Cells(r, colNicht + c * 2).Value)
        vor = Trim(ws.Cells(r, colNicht + c * 2 + 1).Value)

        If nach <> "" And vor <> "" Then

            gefunden = False

            For r2 = 4 To lastRow

                If Trim(ws.Cells(r2, colName).Value) = nach _
                And Trim(ws.Cells(r2, colVorname).Value) = vor Then

                    gefunden = True
                    Exit For

                End If

            Next r2

            If Not gefunden Then
                ws.Cells(r, colNicht + c * 2).Interior.Color = RGB(255, 0, 0)
                ws.Cells(r, colNicht + c * 2 + 1).Interior.Color = RGB(255, 0, 0)
            Else
                ws.Cells(r, colNicht + c * 2).Interior.ColorIndex = xlNone
                ws.Cells(r, colNicht + c * 2 + 1).Interior.ColorIndex = xlNone
            End If

        End If

    Next c

End If

Next r

MsgBox "Prüfung abgeschlossen."

End Sub

Sub ClusterListeErstellen(cluster() As Integer, nameListe() As String)

Dim ws As Worksheet
On Error Resume Next
Set ws = Worksheets("Cluster")
On Error GoTo 0

If ws Is Nothing Then
    Set ws = Worksheets.Add
    ws.Name = "Cluster"
Else
    ws.Cells.Clear
End If

Dim maxCluster As Long
Dim i As Long

'Max Cluster bestimmen
For i = 1 To UBound(cluster)
    If cluster(i) > maxCluster Then maxCluster = cluster(i)
Next i

Dim size() As Long
ReDim size(1 To maxCluster)

'Clustergrößen zählen
For i = 1 To UBound(cluster)
    size(cluster(i)) = size(cluster(i)) + 1
Next i

'Clusterliste erzeugen
Dim order() As Long
ReDim order(1 To maxCluster)

For i = 1 To maxCluster
    order(i) = i
Next i

'Nach Größe sortieren (einfacher Bubble Sort)
Dim j As Long, tmp As Long

For i = 1 To maxCluster - 1
    For j = i + 1 To maxCluster
        If size(order(j)) > size(order(i)) Then
            tmp = order(i)
            order(i) = order(j)
            order(j) = tmp
        End If
    Next j
Next i

'Überschrift
ws.Cells(1, 1) = "Cluster"
ws.Cells(1, 2) = "Größe"
ws.Cells(1, 3) = "Schüler"

Dim rowOut As Long
rowOut = 2

Dim cl As Long
Dim text As String

For i = 1 To maxCluster

    cl = order(i)
    text = ""

    For j = 1 To UBound(cluster)
        If cluster(j) = cl Then
            text = text & nameListe(j) & ", "
        End If
    Next j

    ws.Cells(rowOut, 1) = cl
    ws.Cells(rowOut, 2) = size(cl)
    ws.Cells(rowOut, 3) = Left(text, Len(text) - 2)

    rowOut = rowOut + 1

Next i

ws.Columns.AutoFit

End Sub

Function BerechneZielwerte(geschlecht() As String, _
                          schulform() As String, _
                          rot() As Boolean, _
                          ANZAHL_KLASSEN As Integer) As Variant

    Dim i As Long
    Dim gesM As Long, gesRS As Long, gesRSGy As Long, gesRot As Long
    Dim n As Long
    
    n = UBound(geschlecht)
    
    For i = 1 To n
    
        If geschlecht(i) = "m" Then gesM = gesM + 1
        
        If schulform(i) = "RS" Then gesRS = gesRS + 1
        If schulform(i) = "RSGy" Then gesRSGy = gesRSGy + 1
        
        If rot(i) Then gesRot = gesRot + 1
        
    Next i
    
    Dim zielM As Double, zielRS As Double, zielRSGy As Double, zielRot As Double
    
    zielM = gesM / ANZAHL_KLASSEN
    zielRS = gesRS / ANZAHL_KLASSEN
    zielRSGy = gesRSGy / ANZAHL_KLASSEN
    zielRot = gesRot / ANZAHL_KLASSEN
    
    BerechneZielwerte = Array(zielM, zielRS, zielRSGy, zielRot)

End Function

Function GruppePasstKomplett(grp As Variant, _
                             klasseName As String, _
                             wsA As Worksheet, _
                             colName As Long, _
                             colVorname As Long, _
                             colFix As Long, _
                             colNicht As Long, _
                             lastRow As Long) As Boolean
    '1. gegen bestehende Fix prüfen
    If Not ClusterPasstOhneKonflikt(grp, klasseName, wsA, colName, colVorname, colFix, colNicht, lastRow) Then
        Exit Function
    End If

    '2. intern prüfen (NEU)
    If GruppeHatInterneKonflikte(grp, wsA, colName, colVorname, colNicht, lastRow) Then
        Exit Function
    End If

    GruppePasstKomplett = True

End Function

Function HatKonfliktZwischenZwei(name1 As String, _
                                name2 As String, _
                                wsA As Worksheet, _
                                colName As Long, _
                                colVorname As Long, _
                                colNicht As Long, _
                                lastRow As Long) As Boolean

    Dim r As Long
    Dim fn As String
    Dim nichtText As String
    Dim parts() As String
    Dim p As Variant

    '--- name1 ? name2 prüfen
    For r = 4 To lastRow
    
        fn = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
        
        If NamenMatch(fn, name1) Then
        
            nichtText = wsA.Cells(r, colNicht).Value
            
            If nichtText <> "" Then
            
                parts = Split(nichtText, "|")
                
                For Each p In parts
                    If NamenMatch(CStr(p), name2) Then
                        HatKonfliktZwischenZwei = True
                        Exit Function
                    End If
                Next p
                
            End If
            
            Exit For
            
        End If
        
    Next r

    '--- name2 ? name1 prüfen (wichtig!)
    For r = 4 To lastRow
    
        fn = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
        
        If NamenMatch(fn, name2) Then
        
            nichtText = wsA.Cells(r, colNicht).Value
            
            If nichtText <> "" Then
            
                parts = Split(nichtText, "|")
                
                For Each p In parts
                    If NamenMatch(CStr(p), name1) Then
                        HatKonfliktZwischenZwei = True
                        Exit Function
                    End If
                Next p
                
            End If
            
            Exit For
            
        End If
        
    Next r

End Function



Function GruppeHatInterneKonflikte(grp As Variant, _
                                   wsA As Worksheet, _
                                   colName As Long, _
                                   colVorname As Long, _
                                   colNicht As Long, _
                                   lastRow As Long) As Boolean

    Dim a As Variant, b As Variant

    For Each a In grp
        For Each b In grp
        
            If a <> b Then
            
                If HatKonfliktZwischenZwei(CStr(a), CStr(b), wsA, colName, colVorname, colNicht, lastRow) Then
                    GruppeHatInterneKonflikte = True
                    Exit Function
                End If
                
            End If
            
        Next b
    Next a

End Function



Sub FixiereGroessteCluster()

    Dim wsC As Worksheet, wsA As Worksheet, wsTiefe As Worksheet
    Dim lastRow As Long, r As Long
    
    On Error Resume Next
    Set wsC = Worksheets("MinMaxCluster_V2")
    Set wsA = Worksheets("alle")
    Set wsTiefe = Worksheets("Suchtiefe")
    On Error GoTo 0

    If wsC Is Nothing Or wsA Is Nothing Or wsTiefe Is Nothing Then
        MsgBox "Benötigte Tabellen fehlen!"
        Exit Sub
    End If

    '--- Parameter
    ANZAHL_KLASSEN = wsTiefe.Range("B18").Value
    Dim maxCluster As Long
    maxCluster = wsTiefe.Range("B20").Value
    
    If ANZAHL_KLASSEN < 1 Or ANZAHL_KLASSEN > 7 Then
        MsgBox "Klassenanzahl (B18) muss zwischen 1 und 7 liegen.", vbExclamation
        Exit Sub
    End If
    
    If IsEmpty(wsTiefe.Range("B20").Value) Or wsTiefe.Range("B20").Value = "" Then
        MsgBox "Bitte Anzahl Fixierung Wunschketten in Zelle B20 (Suchtiefe) eintragen.", vbExclamation
        Exit Sub
    End If
    
    If maxCluster < 1 Then
        MsgBox "Anzahl Fixierung Wunschketten (B20) muss mindestens 1 sein.", vbExclamation
        Exit Sub
    End If

    '--- Spalten
    Dim colName As Long, colVorname As Long, colFix As Long, colNicht As Long
    
    colName = FindeSpalteZ2(wsA, "name")
    colVorname = FindeSpalteZ2(wsA, "vorname")
    colFix = FindeSpalteZ2(wsA, "fix")
    colNicht = FindeSpalte(wsA, "besser nicht zusammen mit")

    If colName = 0 Or colVorname = 0 Or colFix = 0 Or colNicht = 0 Then
        MsgBox "Spalten fehlen!"
        Exit Sub
    End If

    lastRow = wsA.Cells(wsA.Rows.Count, colName).End(xlUp).Row

    '====================================================
    ' Klassen-Größen zählen (wichtig!)
    '====================================================
    Dim classSize() As Long
    ReDim classSize(1 To ANZAHL_KLASSEN)

    Dim i As Long
    
    For r = 4 To lastRow
        If wsA.Cells(r, colFix).Value <> "" Then
        
            Dim f As String
            f = wsA.Cells(r, colFix).Value
            
            If Left(f, 1) = "5" Then
                i = Asc(Right(f, 1)) - 96
                If i >= 1 And i <= ANZAHL_KLASSEN Then
                    classSize(i) = classSize(i) + 1
                End If
            End If
            
        End If
    Next r

    '====================================================
    ' Cluster durchgehen
    '====================================================
    Dim k As Long
    
    For k = 1 To maxCluster
    
        Dim rawText As String
        rawText = Trim(CStr(wsC.Cells(k + 1, 3).Value))
        
        If rawText = "" Then Exit For

        '--- Namen aufteilen
        Dim namesArr() As String
        namesArr = Split(rawText, ",")

        Dim cleanedNames As Collection
        Set cleanedNames = New Collection

        Dim x As Variant
        For Each x In namesArr
            If Trim(CStr(x)) <> "" Then
                cleanedNames.Add Trim(CStr(x))
            End If
        Next x

        If cleanedNames.Count = 0 Then GoTo NextCluster

        '--- konfliktfreie Teilgruppen
        Dim teile As Collection
        Set teile = SplitClusterKonfliktfrei(cleanedNames, wsA, colName, colVorname, colNicht, lastRow)

        Dim grp As Variant
        
        For Each grp In teile

            Dim klasseIndex As Long
            Dim klasseName As String
            Dim foundClass As Boolean
            foundClass = False

            '====================================================
            ' BESTE KLASSE SUCHEN (kleinste!)
            '====================================================
            Dim bestClass As Long
            Dim bestSize As Long
            bestSize = 999999
            bestClass = 0

            Dim c As Long
            
            For c = 1 To ANZAHL_KLASSEN
            
                klasseName = "5" & Chr(96 + c)
                
If GruppePasstKomplett(grp, klasseName, wsA, colName, colVorname, colFix, colNicht, lastRow) Then
                    If classSize(c) < bestSize Then
                        bestSize = classSize(c)
                        bestClass = c
                    End If
                    
                End If
                
            Next c

            If bestClass > 0 Then
                klasseIndex = bestClass
                klasseName = "5" & Chr(96 + bestClass)
                foundClass = True
            End If

            If Not foundClass Then GoTo NextGroup

            '====================================================
            ' FIXIEREN
            '====================================================
            Dim n As Variant
            
            For Each n In grp
            
                Dim targetName As String
                targetName = Trim(CStr(n))
                
                If targetName = "" Then GoTo NextName
                
                For r = 4 To lastRow
                
                    Dim fullName As String
                    fullName = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
                    
                    If fullName <> "" Then
                    
                        If NamenMatch(fullName, targetName) Then
                        
                            If wsA.Cells(r, colFix).Value = "" Then
                                wsA.Cells(r, colFix).Value = klasseName
                                classSize(klasseIndex) = classSize(klasseIndex) + 1   ' ? wichtig
                            End If
                            
                            Exit For
                            
                        End If
                        
                    End If
                    
                Next r
                
NextName:
            Next n

NextGroup:
        Next grp

NextCluster:
    Next k

    MsgBox "Cluster stabil und gleichmäßig fixiert!", vbInformation

End Sub


Sub AnalyseUndKonflikteNeu()

    Dim ws As Worksheet
    Set ws = Worksheets("alle")
    
    Dim colL1 As Long
    Dim colName As Long, colVorname As Long
    Dim colM As Long, colW As Long
    Dim colRS As Long, colRSGy As Long
    Dim colWohnort As Long, colFarbe As Long
    Dim colWunsch As Long, colNicht As Long
    Dim colLehrer As Long
    
    '--- Spalten finden
    colName = FindeSpalteZ2(ws, "name")
    colVorname = FindeSpalteZ2(ws, "vorname")
    colM = FindeSpalteZ3(ws, "m")
    colW = FindeSpalteZ3(ws, "w")
    colRS = FindeSpalteZ3(ws, "rs")
    colRSGy = FindeSpalteZ3(ws, "rs/gy")
    colWohnort = FindeSpalteZ3(ws, "wohnort")
    colFarbe = FindeSpalteZ2(ws, "farbe")
    colLehrer = FindeSpalte(ws, "klassenlehrer")
    
    Dim colKommentarMit As Long
    colKommentarMit = FindeSpalte(ws, "kommentar möchte")
    
    colWunsch = colKommentarMit + 1
    colNicht = FindeSpalte(ws, "besser nicht zusammen mit")
    
    colL1 = FindeSpalte(ws, "l1")
    
    If colL1 = 0 Then
        MsgBox "Keine Lösungen (L1–L5) gefunden!", vbExclamation
        Exit Sub
    End If
    
    '--- letzte Zeile
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colName).End(xlUp).Row
    
    Dim n As Long
    n = lastRow - 3
    
    '--- Arrays neu aufbauen (wie im Hauptlauf)
    Dim nameListe() As String
    Dim geschlecht() As String
    Dim schulform() As String
    Dim wohnort() As String
    Dim rot() As Boolean
    Dim wunsch() As String
    Dim nichtMit() As String
    Dim lehrer() As String
    
    ReDim nameListe(1 To n)
    ReDim geschlecht(1 To n)
    ReDim schulform(1 To n)
    ReDim wohnort(1 To n)
    ReDim rot(1 To n)
    ReDim wunsch(1 To n)
    ReDim nichtMit(1 To n)
    ReDim lehrer(1 To n)
    
    Dim i As Long, r As Long
    i = 1
    
    For r = 4 To lastRow
    
        nameListe(i) = NormalisiereName(ws.Cells(r, colName), ws.Cells(r, colVorname))
        
        If colM > 0 And LCase(ws.Cells(r, colM)) = "x" Then geschlecht(i) = "m"
        If colW > 0 And LCase(ws.Cells(r, colW)) = "x" Then geschlecht(i) = "w"
        
        If colRS > 0 And LCase(ws.Cells(r, colRS)) = "x" Then schulform(i) = "RS"
        If colRSGy > 0 And LCase(ws.Cells(r, colRSGy)) = "x" Then schulform(i) = "RSGy"
        
        If colWohnort > 0 Then wohnort(i) = ws.Cells(r, colWohnort)
        If colFarbe > 0 Then rot(i) = (Trim(ws.Cells(r, colFarbe)) <> "")
        If colLehrer > 0 Then lehrer(i) = ws.Cells(r, colLehrer)
        
        '--- Wünsche
        Dim w1 As String, w2 As String
        
        w1 = NormalisiereName(ws.Cells(r, colWunsch), ws.Cells(r, colWunsch + 1))
        w2 = NormalisiereName(ws.Cells(r, colWunsch + 2), ws.Cells(r, colWunsch + 3))
        
        If w1 = "" Then
            wunsch(i) = w2
        ElseIf w2 = "" Then
            wunsch(i) = w1
        Else
            wunsch(i) = w1 & "|" & w2
        End If
        
        '--- Konflikte
        Dim n1 As String, n2 As String
        
        n1 = NormalisiereName(ws.Cells(r, colNicht), ws.Cells(r, colNicht + 1))
        n2 = NormalisiereName(ws.Cells(r, colNicht + 2), ws.Cells(r, colNicht + 3))
        
        If n1 = "" Then
            nichtMit(i) = n2
        ElseIf n2 = "" Then
            nichtMit(i) = n1
        Else
            nichtMit(i) = n1 & "|" & n2
        End If
        
        i = i + 1
        
    Next r
    
    '--- Anzahl Klassen holen
    Dim wsTiefe As Worksheet
    Set wsTiefe = Worksheets("Suchtiefe")
    
    ANZAHL_KLASSEN = wsTiefe.Range("B18").Value
    
    '--- NUR AUSWERTUNG
    Call AnalyseErstellen(ws, nameListe, geschlecht, schulform, rot, wohnort, wunsch, nichtMit, lehrer, colL1, n, ANZAHL_KLASSEN)
    Call WunschFehlerListe(ws, nameListe, wunsch, colL1, n)
    Call KonfliktListe(ws, nameListe, nichtMit, colL1, n)
    
    MsgBox "Analyse und Konflikte aktualisiert!", vbInformation

End Sub








Sub MinMaxCluster_V2(cluster() As Integer, _
                     nichtMit() As String, _
                     nameListe() As String, _
                     nameIndex As Object)
                     
                     
    Dim wsA As Worksheet
    Set wsA = Worksheets("alle")
    
    Dim wsOut As Worksheet
    On Error Resume Next
    Set wsOut = Worksheets("MinMaxCluster_V2")
    On Error GoTo 0
    
    If wsOut Is Nothing Then
        Set wsOut = Worksheets.Add
        wsOut.Name = "MinMaxCluster_V2"
    Else
        wsOut.Cells.Clear
    End If

    '-------------------------------
    ' Spalten
    '-------------------------------
    Dim colName As Long, colVorname As Long
    Dim colKommentarMit As Long, colWunsch As Long
    
    colName = FindeSpalteZ2(wsA, "name")
    colVorname = FindeSpalteZ2(wsA, "vorname")
    colKommentarMit = FindeSpalte(wsA, "kommentar möchte")
    
    If colKommentarMit = 0 Then
        MsgBox "Spalte 'Kommentar möchte' fehlt!"
        Exit Sub
    End If
    
    colWunsch = colKommentarMit + 1

    '-------------------------------
    ' Datenbereich
    '-------------------------------
    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.Count, colName).End(xlUp).Row
    
    Dim n As Long
    n = lastRow - 3

    '-------------------------------
    ' Arrays
    '-------------------------------
    
    Dim wunsch() As String
    
    ReDim nameListe(1 To n)
    ReDim wunsch(1 To n)

    '-------------------------------
    ' Einlesen
    '-------------------------------
    Dim i As Long, r As Long
    i = 1
    
    For r = 4 To lastRow
    
        nameListe(i) = NormalisiereName(wsA.Cells(r, colName), wsA.Cells(r, colVorname))
        
        Dim w1 As String, w2 As String
        
        w1 = NormalisiereName(wsA.Cells(r, colWunsch), wsA.Cells(r, colWunsch + 1))
        w2 = NormalisiereName(wsA.Cells(r, colWunsch + 2), wsA.Cells(r, colWunsch + 3))
        
        If w1 = "" And w2 = "" Then
            wunsch(i) = ""
        ElseIf w1 = "" Then
            wunsch(i) = w2
        ElseIf w2 = "" Then
            wunsch(i) = w1
        Else
            wunsch(i) = w1 & "|" & w2
        End If
        
        i = i + 1
    Next r

    '-------------------------------
    ' Namensindex
    '-------------------------------
   
    Set nameIndex = CreateObject("Scripting.Dictionary")
    
    For i = 1 To n
        nameIndex(nameListe(i)) = i
    Next i

    '-------------------------------
    ' Union-Find
    '-------------------------------
    Dim parent() As Long
    ReDim parent(1 To n)
    
    For i = 1 To n
        parent(i) = i
    Next i

    '-------------------------------
    ' MinMax Logik
    '-------------------------------
    Dim parts() As String, p As Variant
    Dim j As Long
    
    For i = 1 To n
    
        If wunsch(i) <> "" Then
        
            parts = Split(wunsch(i), "|")
            
            Dim bestPartner As Long
            Dim bestMax As Long
            bestMax = 999999
            
            For Each p In parts
            
                If Trim(p) <> "" And nameIndex.Exists(p) Then
                
                    j = nameIndex(p)
                    
                    '--- Simulation
                    Dim parentCopy() As Long
                    parentCopy = parent
                    
                    UF_Union parentCopy, i, j
                    
                    Dim maxSize As Long
                    maxSize = UF_MaxClusterSize(parentCopy)
                    
                    If maxSize < bestMax Then
                        bestMax = maxSize
                        bestPartner = j
                    End If
                    
                End If
                
Next p

' ? gehört INS If wunsch(i)
If bestPartner > 0 Then

    If Not ClusterHatKonflikt_GLOBAL(parent, i, bestPartner, nichtMit, nameListe, nameIndex) Then
        UF_Union parent, i, bestPartner
    End If

End If

End If   ' <-- schließt If wunsch(i)
        
    Next i

    '-------------------------------
    ' Cluster berechnen
    '-------------------------------
    For i = 1 To n
        cluster(i) = UF_Find(parent, i)
    Next i

    '-------------------------------
    ' REPARATUR (WICHTIG!)
    '-------------------------------
    For i = 1 To n
        If cluster(i) < 1 Then cluster(i) = i
    Next i

    '-------------------------------
    ' Neu nummerieren
    '-------------------------------
    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    Dim newID As Long
    newID = 0

    For i = 1 To n
    
        If Not map.Exists(cluster(i)) Then
            newID = newID + 1
            map(cluster(i)) = newID
        End If
        
        cluster(i) = map(cluster(i))
        
    Next i

    '-------------------------------
    ' Größen zählen
    '-------------------------------
    Dim sizeDict As Object
    Set sizeDict = CreateObject("Scripting.Dictionary")
    
    For i = 1 To n
        If sizeDict.Exists(cluster(i)) Then
            sizeDict(cluster(i)) = sizeDict(cluster(i)) + 1
        Else
            sizeDict(cluster(i)) = 1
        End If
    Next i

    '-------------------------------
    ' Sortieren
    '-------------------------------
    Dim keys As Variant
    keys = sizeDict.keys
    
    Dim tmp As Variant
    
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If sizeDict(keys(j)) > sizeDict(keys(i)) Then
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i

    '-------------------------------
    ' Ausgabe
    '-------------------------------
    wsOut.Cells(1, 1) = "Cluster"
    wsOut.Cells(1, 2) = "Größe"
    wsOut.Cells(1, 3) = "Schüler"
    
    Dim rowOut As Long: rowOut = 2
    Dim cl As Variant
    
    For Each cl In keys
    
        Dim txt As String: txt = ""
        
        For i = 1 To n
            If cluster(i) = cl Then
                txt = txt & nameListe(i) & ", "
            End If
        Next i
        
        wsOut.Cells(rowOut, 1) = cl
        wsOut.Cells(rowOut, 2) = sizeDict(cl)
        wsOut.Cells(rowOut, 3) = Left(txt, Len(txt) - 2)
        
        rowOut = rowOut + 1
        
    Next cl

    wsOut.Columns.AutoFit
    
    MsgBox "MinMaxCluster V2 stabil erstellt!", vbInformation

End Sub



Sub UF_Union(parent() As Long, ByVal a As Long, ByVal b As Long)

    Dim ra As Long, rb As Long
    
    ra = UF_Find(parent, a)
    rb = UF_Find(parent, b)
    
    If ra <> rb Then
        parent(rb) = ra
    End If

End Sub

Function ClusterHatKonflikt_GLOBAL(parent() As Long, _
                                   ByVal a As Long, _
                                   ByVal b As Long, _
                                   ByRef nichtMit() As String, _
                                   ByRef nameListe() As String, _
                                   ByVal nameIndex As Object) As Boolean

    Dim ra As Long, rb As Long
    ra = UF_Find(parent, a)
    rb = UF_Find(parent, b)

    Dim i As Long, j As Long
    Dim parts() As String, p As Variant

    For i = 1 To UBound(parent)
    
        If UF_Find(parent, i) = ra Then
        
            If nichtMit(i) <> "" Then
            
                parts = Split(nichtMit(i), "|")
                
                For Each p In parts
                
                    If nameIndex.Exists(p) Then
                    
                        j = nameIndex(p)
                        
                        If UF_Find(parent, j) = rb Then
                            ClusterHatKonflikt_GLOBAL = True
                            Exit Function
                        End If
                        
                    End If
                    
                Next p
                
            End If
            
        End If
        
    Next i

End Function

Function UF_Find(parent() As Long, ByVal x As Long) As Long

    If parent(x) <> x Then
        parent(x) = UF_Find(parent, parent(x))
    End If
    
    UF_Find = parent(x)

End Function

Function UF_MaxClusterSize(parent() As Long) As Long

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, r As Long
    
    For i = 1 To UBound(parent)
    
        r = UF_Find(parent, i)
        
        If dict.Exists(r) Then
            dict(r) = dict(r) + 1
        Else
            dict(r) = 1
        End If
        
        If dict(r) > UF_MaxClusterSize Then
            UF_MaxClusterSize = dict(r)
        End If
        
    Next i

End Function

Function ClusterPasstOhneKonflikt(names As Variant, _
                                  klasseName As String, _
                                  wsA As Worksheet, _
                                  colName As Long, _
                                  colVorname As Long, _
                                  colFix As Long, _
                                  colNicht As Long, _
                                  lastRow As Long) As Boolean

Dim n As Variant, r As Long
Dim fullName As String
Dim nichtText As String
Dim parts() As String
Dim p As Variant
Dim r2 As Long

'--- Sicherheit: names muss Array sein
If IsEmpty(names) Then
    ClusterPasstOhneKonflikt = True
    Exit Function
End If




For Each n In names

    fullName = Trim(CStr(n))
    
    
    'Schülerzeile finden
    For r = 4 To lastRow
    
        Dim fn As String
        fn = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
        
        If NamenMatch(fn, fullName) Then
        
            nichtText = wsA.Cells(r, colNicht).Value
            
            If nichtText <> "" Then
            
                parts = Split(nichtText, "|")
                
                For Each p In parts
                
                    For r2 = 4 To lastRow
                    
                        Dim fn2 As String
                        fn2 = Trim(wsA.Cells(r2, colName).Value & " " & wsA.Cells(r2, colVorname).Value)
                        
                        If NamenMatch(fn2, CStr(p)) Then
                        
                            'ist der Konfliktpartner schon in der Zielklasse?
                            If wsA.Cells(r2, colFix).Value = klasseName Then
                                ClusterPasstOhneKonflikt = False
                                Exit Function
                            End If
                            
                        End If
                        
                    Next r2
                    
                Next p
                
            End If
            
            Exit For
            
        End If
        
    Next r

Next n

ClusterPasstOhneKonflikt = True

End Function

Function SplitClusterKonfliktfrei(names As Variant, _
                                 wsA As Worksheet, _
                                 colName As Long, _
                                 colVorname As Long, _
                                 colNicht As Long, _
                                 lastRow As Long) As Collection

    Dim result As New Collection
    
    Dim name1 As Variant
    Dim grp As Collection
    Dim added As Boolean

    '--- Prüfen: ist es eine Collection?
    If TypeName(names) = "Collection" Then
    
        For Each name1 In names
        
            name1 = Trim(CStr(name1))
            If name1 = "" Then GoTo NextName
            
            added = False
            
            '--- vorhandene Gruppen testen
            For Each grp In result
            
                If Not HatKonfliktMitGruppe(CStr(name1), grp, wsA, colName, colVorname, colNicht, lastRow) Then
                    grp.Add name1
                    added = True
                    Exit For
                End If
                
            Next grp
            
            '--- neue Gruppe
            If Not added Then
                Dim newGrp As New Collection
                newGrp.Add name1
                result.Add newGrp
            End If

NextName:
        Next name1
        
    Else
        'Fallback (falls doch Array)
        Dim i As Long
        For i = LBound(names) To UBound(names)
        
            name1 = Trim(CStr(names(i)))
            If name1 = "" Then GoTo NextArray
            
            added = False
            
            For Each grp In result
            
                If Not HatKonfliktMitGruppe(name1, grp, wsA, colName, colVorname, colNicht, lastRow) Then
                    grp.Add name1
                    added = True
                    Exit For
                End If
                
            Next grp
            
            If Not added Then
                Dim newGrp2 As New Collection
                newGrp2.Add name1
                result.Add newGrp2
            End If

NextArray:
        Next i
    End If

    Set SplitClusterKonfliktfrei = result

End Function

Function HatKonfliktMitGruppe(ByVal name1 As Variant, _
                              ByVal grp As Variant, _
                              wsA As Worksheet, _
                              colName As Long, _
                              colVorname As Long, _
                              colNicht As Long, _
                              lastRow As Long) As Boolean

    Dim r As Long, fn As String
    Dim nichtText As String
    Dim parts() As String, p As Variant
    Dim member As Variant
    Dim n1 As String
    
    n1 = Trim(CStr(name1))

    '========================================
    ' 1. name1 ? Gruppe prüfen
    '========================================
    For r = 4 To lastRow
    
        fn = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
        
        If NamenMatch(fn, n1) Then
        
            nichtText = wsA.Cells(r, colNicht).Value
            
            If nichtText <> "" Then
            
                parts = Split(nichtText, "|")
                
                For Each p In parts
                    For Each member In grp
                        If NamenMatch(CStr(p), CStr(member)) Then
                            HatKonfliktMitGruppe = True
                            Exit Function
                        End If
                    Next member
                Next p
                
            End If
            
            Exit For
            
        End If
        
    Next r

    '========================================
    ' 2. Gruppe ? name1 prüfen (NEU!!)
    '========================================
    For Each member In grp
    
        For r = 4 To lastRow
        
            fn = Trim(wsA.Cells(r, colName).Value & " " & wsA.Cells(r, colVorname).Value)
            
            If NamenMatch(fn, CStr(member)) Then
            
                nichtText = wsA.Cells(r, colNicht).Value
                
                If nichtText <> "" Then
                
                    parts = Split(nichtText, "|")
                    
                    For Each p In parts
                        If NamenMatch(CStr(p), n1) Then
                            HatKonfliktMitGruppe = True
                            Exit Function
                        End If
                    Next p
                    
                End If
                
                Exit For
                
            End If
            
        Next r
        
    Next member

End Function
