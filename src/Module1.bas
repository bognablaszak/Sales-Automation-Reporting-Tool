Attribute VB_Name = "Module1"
Sub Importuj_dane()
Separator = "-" 'Do odzielenia kolumny Województwo-miasto od siebie
'Wy³¹czy³am niepotrzebne komunikaty i zdarzenia do przyspieszenia dzia³ania
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.EnableEvents = False
'Zmienne deklaracje
Dim folder As String
Dim plik As String
Dim ws As Worksheet
Dim wbTXT As Workbook
Dim lastRow As Long
Dim rngSource As Range
    
folder = InputBox("Podaj œcie¿kê do folderu z plikami txt:", "Import danych") 'Zapytanie do u¿ytkownika o œcie¿kê
If folder = "" Then GoTo Koniec 'Jeœli œcie¿ka zostanie pusta to idzie na koniec
If Right(folder, 1) <> "\" Then folder = folder & "\"
'Przygotowywanie nowego arkusza Dane
On Error Resume Next
Set ws = Worksheets("Dane") 'Ustawiamy na arkusz Dane
On Error GoTo 0
If ws Is Nothing Then 'Jeœli nie ma jeszcze arkusza Dane to dodaje nowy
    Set ws = Worksheets.Add
    ws.Name = "Dane"
Else 'Jak taki arkusz istnieje to czyœci go (usuwa tabelê i dane)
If ws.ListObjects.count > 0 Then ws.ListObjects(1).Delete
ws.Cells.Clear
End If

ws.Range("A1:E1") = Array("Brand", "Produkt", "Tydzien", "Sprzedaz", "Wojewodztwo-Miasto")
lastRow = 2 'tworzenie nag³ówków i ustawienie, aby plik zosta³ zaimportowany od 2 wiersza w arkuszu

plik = Dir(folder & "*.txt") 'pierwszy plik txt w folderze
If plik = "" Then 'Jak nie ma takiego pliku to idzie na koniec
    MsgBox "Nie znaleziono plików .txt w podanym folderze.", vbExclamation
    GoTo Koniec
End If
    
Do While plik <> "" 'Jeœli zmienna plik nie jest pusta
Set wbTXT = Workbooks.Open(folder & plik)   'otwieramy txt
Set rngSource = wbTXT.Sheets(1).Range("A1").CurrentRegion
If rngSource.Rows.count > 1 Then    'jak w pliku jest cos wiecej ni¿ nag³ówek to pomija nag³ówek i reszte wkleja do arkusza
rngSource.Offset(1).Resize(rngSource.Rows.count - 1).Copy ws.Cells(lastRow, 1)
lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1 'oblicza ktory wiersz aktualnie jest ostatni
End If
        
wbTXT.Close False 'zamyka bez zapisu
plik = Dir()  'przechodzimy do kolejnego pliku
Loop   'na pocz¹tek pêtli wraca
    
If lastRow <= 2 Then    'jak s¹ mniej ni¿ 2 wiersze to nie zaimportowano danych i idziemy na koniec
    MsgBox "Nie zaimportowano ¿adnych danych.", vbExclamation
    GoTo Koniec
End If
'Sortowanie po kolumnnie Tydzien
With ws.Sort
    .SortFields.Clear
    .SortFields.Add Key:=ws.Range("C2:C" & lastRow - 1), _
    SortOn:=xlSortOnValues, Order:=xlAscending
    .SetRange ws.Range("A1:E" & lastRow - 1)
    .Header = xlYes
    .Apply
End With

Dim i As Long
Dim pelnyTekst As String
Dim pos As Integer
'Rozdzielenie kolumny Wojewodztwo-Miasto (napisane z pomoca AI, gdy¿ pojawi³y siê przeszkody w trakcie)
For i = 2 To lastRow - 1
pelnyTekst = ws.Cells(i, "E").Value 'w kolumnie E jest Wojewodztwo-Miasto
pos = InStr(1, pelnyTekst, "-") 'szuka pozycji myslnik
If pos > 0 Then
ws.Cells(i, "F").Value = Trim(Left(pelnyTekst, pos - 1)) 'lewa strona myslnika do kolumny F
ws.Cells(i, "G").Value = Trim(Mid(pelnyTekst, pos + 1)) 'prawa strona myslnika do kolumny G
Else
ws.Cells(i, "F").Value = pelnyTekst 'jak nie ma myslnika to calosc do kolumny F
End If
Next i

ws.Columns("E").Delete Shift:=xlToLeft 'usuwamy kolumne Wojewodztwo-Miasto i resztê przesuwamy w lewo
ws.Range("E1").Value = "Województwo"
ws.Range("F1").Value = "Miasto"
'Projektowanie tabeli
Dim lo As ListObject
Dim lastRowData As Long
lastRowData = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F" & lastRowData), , xlYes)
lo.Name = "TabelaDanych"
lo.TableStyle = "TableStyleLight20"
    MsgBox "Dane zosta³y pomyœlnie zaimportowane!", vbInformation
Koniec:
'Przywracanie mo¿liwoœci wyskakiwania wczeœniej wy³¹czonych okienek i zdarzeñ
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True

End Sub
Sub Generuj_pdf()
'Zmienne
Dim wsRaport As Worksheet
Dim wsApp As Worksheet
Dim pt As PivotTable
Dim kategoria As String
Dim i As Long
Dim sciezka As String
Dim nazwaPliku As String
Dim element As String
Set wsRaport = Worksheets("Raport")
Set wsApp = Worksheets("Aplikacja")
On Error Resume Next
Set pt = wsRaport.PivotTables("PivotGlowny") 'Tabela przestawna
On Error GoTo 0
If pt Is Nothing Then 'jak tabeli nie ma to przerwanie
    MsgBox "Najpierw musisz odœwie¿yæ raport!", vbExclamation
    Exit Sub
End If
    
kategoria = wsApp.Range("ChosenColumn").Value  'pobiera wartoœæ wybran¹ przez uzytkownika (wysuwana tabelka)
sciezka = ThisWorkbook.Path & "\"  'sciezka gdzie zapisze siê pdf

Dim decyzja As Integer
decyzja = MsgBox("Czy chcesz wygenerowaæ raporty PDF dla WSZYSTKICH pozycji?" & vbNewLine & vbNewLine & _
        "TAK - Generuj automatycznie dla ca³ej listy" & vbNewLine & _
        "NIE - Pozwól mi wpisaæ jedn¹ konkretn¹ nazwê", vbYesNoCancel + vbQuestion, "Wybór generowania")
If decyzja = vbCancel Then Exit Sub

Application.ScreenUpdating = False 'ekran nie miga
'Generowanie ca³ej listy
If decyzja = vbYes Then
For i = 1 To pt.DataBodyRange.Rows.count 'od 1 do liczby wierszy w tabeli
Call ZapiszJedenPDF(pt, i, kategoria, sciezka, wsRaport) 'przywo³anie innego makra
Next i
    MsgBox "Gotowe! Wszystkie pliki PDF zapisane w: " & sciezka, vbInformation
Else

Dim wybor As String
wybor = InputBox("Wpisz nazwê elementu (np. Warszawa):", "Generuj jeden PDF")
If wybor = "" Then
Application.ScreenUpdating = True 'w³¹czenie z powrotem, bo nastapi koniec makra
Exit Sub
End If
'porównywanie tekstu z tym co wpisa³ u¿ytkownik + zabezpieczenie w razie gdyby u¿ytkownik wpisa³ z ma³ej litery albo doda³ spacjê
Dim znalezionyIndeks As Long
znalezionyIndeks = 0
Dim nazwaZTabeli As String
Dim szukanaNazwa As String
szukanaNazwa = Trim(LCase(wybor))
For i = 1 To pt.DataBodyRange.Rows.count    'nazwa miasta z tabeli
nazwaZTabeli = pt.DataBodyRange.Cells(i, 1).Offset(0, -1).Value 'miasta s¹ o jedn¹ kolumnê w lewo od liczb
If Trim(LCase(nazwaZTabeli)) = szukanaNazwa Then
    znalezionyIndeks = i
Exit For
End If
Next i
If znalezionyIndeks > 0 Then
    Call ZapiszJedenPDF(pt, znalezionyIndeks, kategoria, sciezka, wsRaport)
    MsgBox "Zapisano plik PDF dla: " & wybor, vbInformation
Else
    MsgBox "Nie znaleziono nazwy: '" & wybor & "'" & vbNewLine & _
    "Upewnij siê, ¿e wpisujesz nazwê dok³adnie tak, jak widaæ j¹ w tabeli Excela.", vbExclamation
End If
End If
Call Uzupelnij_Statystyki(pt, 1) 'przywraca z powrotem raport do stanu poczatkowego
Application.ScreenUpdating = True 'w³¹czamy z powrotem
End Sub
Sub ZapiszJedenPDF(pt As PivotTable, idx As Long, kat As String, sciezka As String, ws As Worksheet)
'to makro sluzy tylko do Generuj_pdf, wiec zamiast zmiennych zastosowa³am argumenty zeby dwa razy nie pisac
Dim el As String
Dim plik As String
Call Uzupelnij_Statystyki(pt, idx)
'dostosowanie tresci do kartki A4, skalowanie itd.
With ws.PageSetup
.Zoom = False
.FitToPagesWide = 1
.FitToPagesTall = 1
.Orientation = xlPortrait 'Pionowo
.CenterHorizontally = True
.LeftMargin = Application.InchesToPoints(0.2)
.RightMargin = Application.InchesToPoints(0.2)
.TopMargin = Application.InchesToPoints(0.2)
.BottomMargin = Application.InchesToPoints(0.2)
End With
'generowanie nazwy pliku
el = pt.DataBodyRange.Cells(idx, 1).Offset(0, -1).Value
plik = "Raport_" & UCase(kat) & "_" & el & ".pdf"
On Error Resume Next
'w³aœciwoœci zapisywanego pliku
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=sciezka & plik, _
Quality:=xlQualityStandard, IncludeDocProperties:=True, _
IgnorePrintAreas:=False, OpenAfterPublish:=False
On Error GoTo 0
End Sub
Sub Rysuj_Tabelke(pt As PivotTable)
Dim wsRaport As Worksheet
Dim i As Long
Dim IloscDanych As Long
Dim SumaCalkowita As Double
Dim limit As Long
Const WierszStart As Long = 38  'wiersz startowy
Const Kol_Lp As String = "C"       'kolumna na pozycji
Const Kol_Nazwa As String = "D"    'kolumna dla produktu
Const Kol_Wartosc As String = "F"  'kolumna dla wartosci w z³
Const Kol_Udzial As String = "I"   'kolumna dla udzia³u w %
Set wsRaport = Worksheets("Raport")
'czyszczenie tresci
wsRaport.Range(Kol_Lp & WierszStart & ":" & Kol_Lp & (WierszStart + 5)).ClearContents
wsRaport.Range(Kol_Nazwa & WierszStart & ":" & Kol_Nazwa & (WierszStart + 5)).ClearContents
wsRaport.Range(Kol_Wartosc & WierszStart & ":" & Kol_Wartosc & (WierszStart + 5)).ClearContents
wsRaport.Range(Kol_Udzial & WierszStart & ":" & Kol_Udzial & (WierszStart + 5)).ClearContents
If pt Is Nothing Then Exit Sub
IloscDanych = pt.DataBodyRange.Rows.count
SumaCalkowita = Application.WorksheetFunction.Sum(pt.DataBodyRange) 'suma sprzedazy
If IloscDanych < 5 Then limit = IloscDanych Else limit = 5  'wybiera top 5, chyba ¿e w danych nie ma piêciu dostêpnych to tyle, ile jest
'Wstawianie danych do tabeli
For i = 1 To limit
wsRaport.Range(Kol_Lp & (WierszStart + i - 1)).Value = i
wsRaport.Range(Kol_Nazwa & (WierszStart + i - 1)).Value = pt.DataBodyRange.Cells(i, 1).Offset(0, -1).Value
wsRaport.Range(Kol_Wartosc & (WierszStart + i - 1)).Value = pt.DataBodyRange.Cells(i, 1).Value
wsRaport.Range(Kol_Udzial & (WierszStart + i - 1)).Value = pt.DataBodyRange.Cells(i, 1).Value / SumaCalkowita
Next i

Dim zakresFormatowania As Range
'formatowanie grupowe (aby nie pisaæ tego samego kilka razy)
Set zakresFormatowania = Application.Union( _
wsRaport.Range(Kol_Lp & WierszStart & ":" & Kol_Lp & (WierszStart + 4)), _
wsRaport.Range(Kol_Nazwa & WierszStart & ":" & Kol_Nazwa & (WierszStart + 4)), _
wsRaport.Range(Kol_Wartosc & WierszStart & ":" & Kol_Wartosc & (WierszStart + 4)), _
wsRaport.Range(Kol_Udzial & WierszStart & ":" & Kol_Udzial & (WierszStart + 4)))
With zakresFormatowania
    .Font.Color = RGB(255, 255, 255)
    .Font.Bold = True
    .Font.Size = 14
    .VerticalAlignment = xlCenter
End With
'dodanie z³otówek i procentów
wsRaport.Range(Kol_Lp & WierszStart & ":" & Kol_Lp & (WierszStart + 4)).HorizontalAlignment = xlCenter
With wsRaport.Range(Kol_Nazwa & WierszStart & ":" & Kol_Nazwa & (WierszStart + 4))
    .NumberFormat = "General"
    .HorizontalAlignment = xlLeft
End With

With wsRaport.Range(Kol_Wartosc & WierszStart & ":" & Kol_Wartosc & (WierszStart + 4))
    .NumberFormat = "# ##0 z³"
    .HorizontalAlignment = xlCenter
End With

With wsRaport.Range(Kol_Udzial & WierszStart & ":" & Kol_Udzial & (WierszStart + 4))
    .NumberFormat = "0.00%"
    .HorizontalAlignment = xlCenter
End With
End Sub

