Attribute VB_Name = "Module2"
Sub Odswiez_raport()
'zmienne
Dim wsDane As Worksheet
Dim wsRaport As Worksheet
Dim wsApp As Worksheet
Dim pc As PivotCache
Dim pt As PivotTable
Dim ch As ChartObject
Dim kategoria As String
Application.ScreenUpdating = False 'wy³¹czenie migania ekranu
Set wsDane = Worksheets("Dane")
Set wsRaport = Worksheets("Raport")
Set wsApp = Worksheets("Aplikacja")
On Error Resume Next
kategoria = wsApp.Range("ChosenColumn").Value 'sprawdzenie, co u¿ytkownik wybra³ jako g³ówn¹ kolumnê
On Error GoTo 0
'jeœli u¿ytkownik nie zaimportuje danych
If wsDane.ListObjects.count = 0 Then
    MsgBox "Brak danych! Kliknij 'Importuj Dane'.", vbExclamation
    Exit Sub
End If
'usuwanie starego wykresu
For Each ch In wsRaport.ChartObjects
ch.Delete
Next ch
'wyczyszczenie schowanych komórek i tabeli
wsRaport.Columns("Z:AZ").Hidden = False 'ods³oniêcie
Dim ptOld As PivotTable
For Each ptOld In wsRaport.PivotTables
ptOld.TableRange2.Clear
Next ptOld
'czyszczenie komórek
wsRaport.Range("Z:AK").ClearContents
wsRaport.Range("M:N").ClearContents
'zapisanie danych w pamieci komputera i stworzenie tabeli
Set pc = ActiveWorkbook.PivotCaches.Create( _
SourceType:=xlDatabase, _
SourceData:=wsDane.ListObjects("TabelaDanych"))
Set pt = pc.CreatePivotTable(TableDestination:=wsRaport.Range("Z1"), TableName:="PivotGlowny")
'konfigurowanie tabeli
With pt
    .AddDataField .PivotFields("Sprzedaz"), "Suma Sprzeda¿y", xlSum
    .PivotFields("Suma Sprzeda¿y").NumberFormat = "# ##0 z³"
On Error Resume Next
    .PivotFields(kategoria).Orientation = xlRowField
On Error GoTo 0
    .PivotFields(kategoria).AutoSort xlDescending, "Suma Sprzeda¿y"
End With
'Wykres s³upkowy
Dim rngWykres As Range
Set rngWykres = wsRaport.Range("B21:L35")
Dim newChart As ChartObject
Set newChart = wsRaport.ChartObjects.Add(Left:=rngWykres.Left, Top:=rngWykres.Top, Width:=rngWykres.Width, Height:=rngWykres.Height)
newChart.Name = "WykresGlowny"
With newChart.Chart
    .SetSourceData Source:=pt.TableRange1
    .ChartType = xlColumnClustered
    .HasTitle = True
    .ChartTitle.Text = "ANALIZA SPRZEDA¯Y WG: " & UCase(kategoria)
    .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 105, 180)
    .HasLegend = False
End With

Call Uzupelnij_Statystyki(pt, 1)
Call Rysuj_Tabelke(pt)
wsRaport.Columns("Z:AZ").Hidden = True  'ukrycie kolumny
wsRaport.Activate
wsRaport.Range("A1").Select
Application.ScreenUpdating = True
'data aktualizacji
    MsgBox "Raport odœwie¿ony!", vbInformation
Worksheets("Raport").Range("B3").Value = "Ostatnia aktualizacja: " & Format(Now, "dd.mm.yyyy, hh:mm")
End Sub
Sub Uzupelnij_Statystyki(pt As PivotTable, indeks As Long)
'tutaj liczy siê wszystkie matematyczne procedury typu œrednia, mediana itd.
Dim wsRaport As Worksheet
Dim wsApp As Worksheet
Set wsRaport = Worksheets("Raport")
Set wsApp = Worksheets("Aplikacja")
Dim WybranaNazwa As String, WybranaWartosc As Double
Dim LiderNazwa As String, LiderWartosc As Double
Dim OstatniNazwa As String, OstatniWartosc As Double
Dim SumaCalkowita As Double, Srednia As Double, Mediana As Double
Dim i As Long
Dim licznik As Long
Dim sumaDoSredniej As Double
Dim nazwaWiersza As String
'jakby okazalo sie, ¿e jest b³¹d w tabeli przestawnej
If pt Is Nothing Then Exit Sub
If indeks > pt.DataBodyRange.Rows.count Then Exit Sub
On Error Resume Next
'ustawienie domyœlnie na zero
sumaDoSredniej = 0
licznik = 0
SumaCalkowita = 0
    
For i = 1 To pt.DataBodyRange.Rows.count
nazwaWiersza = UCase(pt.DataBodyRange.Cells(i, 1).Offset(0, -1).Value) 'pobieranie nazwy (komórka z kwota, przesuniêcie w lewo, zamiana tekstu na du¿e litery dla porównywania)
'Mia³am problem z liczeniem œredniej i by³a zawy¿ona, wiêc tutaj odfiltrowuje siê komórka sumuj¹ca
If nazwaWiersza <> "SUMA KOÑCOWA" Then
Dim wartoscWiersza As Double
'sumowanie
wartoscWiersza = pt.DataBodyRange.Cells(i, 1).Value
sumaDoSredniej = sumaDoSredniej + wartoscWiersza
licznik = licznik + 1
SumaCalkowita = SumaCalkowita + wartoscWiersza 'Sumujemy tylko elementy
End If
Next i
'srednia
If licznik > 0 Then
Srednia = sumaDoSredniej / licznik
Else
Srednia = 0
End If
'mediana
Mediana = Application.WorksheetFunction.Median(pt.DataBodyRange)
'szukamy lidera czyli pierwszy wiersz
LiderNazwa = pt.DataBodyRange.Cells(1, 1).Offset(0, -1).Value
LiderWartosc = pt.DataBodyRange.Cells(1, 1).Value
'Sprawdzenie czy ostatni wiersz jest sum¹ koñcow¹, jeœli tak to idzie o wiersz do góry (¿eby raport nie pokaza³ najmniejsza sprzeda¿ jako ta suma)
Dim idxOstatni As Long
idxOstatni = pt.DataBodyRange.Rows.count
Do While UCase(pt.DataBodyRange.Cells(idxOstatni, 1).Offset(0, -1).Value) = "SUMA KOÑCOWA" _
And idxOstatni > 1
idxOstatni = idxOstatni - 1
Loop
OstatniNazwa = pt.DataBodyRange.Cells(idxOstatni, 1).Offset(0, -1).Value
OstatniWartosc = pt.DataBodyRange.Cells(idxOstatni, 1).Value
'to, co wybra³ u¿ytkownik
WybranaNazwa = pt.DataBodyRange.Cells(indeks, 1).Offset(0, -1).Value
WybranaWartosc = pt.DataBodyRange.Cells(indeks, 1).Value
On Error GoTo 0
'wpisywanie do arkusza (makro Wpisz jest ni¿ej)
Wpisz "£¹czna sprzeda¿ we wszystkich", Format(SumaCalkowita, "# ##0 z³")
Wpisz "Przeciêtna sprzeda¿", Format(Srednia, "# ##0 z³")
Wpisz "Mediana sprzeda¿y", Format(Mediana, "# ##0 z³")
Wpisz "Najwiêksza sprzeda¿", LiderNazwa & "   " & Format(LiderWartosc, "# ##0 z³")
Wpisz "Najmniejsza sprzeda¿", OstatniNazwa & "   " & Format(OstatniWartosc, "# ##0 z³")

Dim naglowek As Range
Dim kat As String
kat = wsApp.Range("ChosenColumn").Value
'nag³ówek
Set naglowek = wsRaport.Cells.Find("RAPORT DLA", LookAt:=xlPart)
If Not naglowek Is Nothing Then
    naglowek.Value = "RAPORT DLA: " & UCase(kat) & " - " & UCase(WybranaNazwa)
End If
Wpisz "£¹czna wielkoœæ sprzeda¿y", Format(WybranaWartosc, "# ##0 z³")
'Zabezpieczenie przed dzieleniem przez zero
If SumaCalkowita > 0 Then
    Wpisz "Udzia³ w ca³kowitej sprzeda¿y", Format(WybranaWartosc / SumaCalkowita, "0.00%")
Else
    Wpisz "Udzia³ w ca³kowitej sprzeda¿y", "0%"
End If
Wpisz "Pozycja", CStr(indeks)

Dim celOcena As Range
'ocena
Set celOcena = wsRaport.Cells.Find("Ocena sprzeda¿y:", LookAt:=xlPart)
If Not celOcena Is Nothing Then
'Czyszczenie
celOcena.Offset(1, 0).ClearContents
celOcena.Offset(0, 2).ClearContents
celOcena.Offset(0, 3).ClearContents
celOcena.Offset(0, 6).ClearContents
celOcena.Offset(0, 6).ClearContents
'warunki dla oceny
If WybranaWartosc >= Srednia Then
celOcena.Offset(0, 6).Value = "Gratulujemy bardzo dobrego wyniku!"
celOcena.Offset(0, 6).Font.Color = RGB(34, 139, 34)
Else
celOcena.Offset(0, 6).Value = "Wielkoœæ sprzeda¿y poni¿ej œredniej."
celOcena.Offset(0, 6).Font.Color = RGB(255, 0, 0)
End If
celOcena.Offset(0, 6).Font.Bold = True
celOcena.Offset(0, 6).HorizontalAlignment = xlLeft
End If
'wykres+kolorowanie wybranego s³upka
On Error Resume Next
Dim ch As Chart
Set ch = wsRaport.ChartObjects("WykresGlowny").Chart
Dim s As Series
Set s = ch.SeriesCollection(1)
s.Format.Fill.ForeColor.RGB = RGB(53, 130, 196)
If indeks <= s.Points.count Then
s.Points(indeks).Format.Fill.ForeColor.RGB = RGB(200, 50, 100)
End If
On Error GoTo 0
End Sub
Sub Wpisz(Tekst As String, wartosc As String)
Dim c As Range
'przeszukiwanie arkusza raport
Set c = Worksheets("Raport").Cells.Find(What:=Tekst, LookIn:=xlValues, LookAt:=xlPart)
If Not c Is Nothing Then
'czyszczenie
Worksheets("Raport").Range(c.Offset(0, 1), c.Offset(0, 5)).ClearContents
'wpisywanie i formatowanie
With c.Offset(0, 6)
    .Value = wartosc
    .HorizontalAlignment = xlLeft
    .Font.Bold = True
End With
End If
End Sub
