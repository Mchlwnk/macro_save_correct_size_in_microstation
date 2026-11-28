Attribute VB_Name = "PrintCorrectSize"
'==============================================================================
' PrintCorrectSize
'
' Makro VBA dla MicroStation V8i.
'
' Zapisuje kazdy model do druku z aktywnego pliku DGN do osobnego PDF-a
' i ustawia rozmiar strony PDF na rozmiar tego modelu. Rozwiazuje typowy
' problem drukowania zbiorczego: przy jednym rozmiarze papieru arkusze A4
' sa pomniejszone, a wieksze arkusze sa obcinane.
'
' Rozmiar strony jest wykrywany w tej kolejnosci:
'   1. definicja arkusza w sheet modelu (Model Properties > Sheet > Size)
'   2. elementy na levelu ramki (nazwa zawiera jedna z BORDER_LEVEL_HINTS)
'   3. zaznaczenie w modelu (jesli cos jest zaznaczone) -> rozwiazanie awaryjne
'   4. najwiekszy zamkniety element w modelu
'   5. granice (extents) wszystkich elementow graficznych w modelu
'
' Zmierzona strona w granicach STANDARD_SIZE_TOLERANCE_MM od A4 albo A3 jest
' zaokraglana do dokladnego rozmiaru ISO. Pozostale rozmiary zachowuja
' zmierzona szerokosc i wysokosc.
'
' Makro korzysta z wbudowanego sterownika Bentley PDF (pdf.pltcfg) i z
' udokumentowanych key-inow PRINT ..., wiec nie potrzeba zewnetrznego
' PDF Creatora.
'
' Cel: MicroStation V8i (SELECTseries), VBA 6/7.
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' Ustawienia do edycji przez uzytkownika
'------------------------------------------------------------------------------

' Sterownik Bentley PDF. Plik jest szukany przez MS_PLTCFG_PATH.
Private Const PDF_DRIVER As String = "pdf.pltcfg"

' Sprawdz, czy PDF_DRIVER istnieje, zanim makro zacznie dzialac na rysunku.
' Ustaw False, jesli sterownik jest znajdowany przez sciezke spoza
' MS_PLTCFG_PATH.
Private Const CHECK_PDF_DRIVER As Boolean = True

' Podfolder tworzony obok pliku DGN. Ustaw "", aby zapisywac obok DGN.
Private Const OUTPUT_SUBFOLDER As String = "PDF"

' Fragmenty nazw leveli ramki. Wielkosc liter nie ma znaczenia.
' Oddzielaj przecinkami. Ustaw "", aby pominac ten etap wykrywania.
' "RAMKA" pasuje do "M_RAMKA" i " _RAMKA".
Private Const BORDER_LEVEL_HINTS As String = "RAMKA"

' Jesli True, zaznaczenie jest uzywane awaryjnie, gdy nie ma ramki.
' Definicja sheet modelu i level ramki zawsze maja pierwszenstwo.
Private Const USE_SELECTION_AS_PAGE As Boolean = True

' Jesli DGN zawiera sheet modele, drukuj tylko je. Modele rysunkowe w takim
' pliku to zwykle geometria zrodlowa, a nie arkusze do wydruku.
' Domyslnie False, bo to makro jest przeznaczone dla modeli rysunkowych,
' z ktorych kazdy ma wlasna ramke.
Private Const PRINT_ONLY_SHEETS_IF_PRESENT As Boolean = False

' Rozmiary standardowe, ktore sa zaokraglane do dokladnego wymiaru ISO.
' Tylko te nazwy sa zaokraglane; pozostale rozmiary zachowuja zmierzona
' szerokosc i wysokosc. Domyslnie A4 i A3 z tolerancja 5 mm.
Private Const STANDARD_SIZES_TO_SNAP As String = "A4,A3"
Private Const STANDARD_SIZE_TOLERANCE_MM As Double = 5#

' Rozsadne granice wykrytej strony (milimetry).
Private Const MIN_PAGE_MM As Double = 40#
Private Const MAX_PAGE_MM As Double = 10000#

' Ile sekund czekac na jeden PDF, zanim makro uzna, ze sie nie udalo.
Private Const PDF_TIMEOUT_SEC As Long = 180

' Zapisz krotki log obok PDF-ow.
Private Const WRITE_LOG As Boolean = True

'------------------------------------------------------------------------------
' Typy prywatne i stan modulu
'------------------------------------------------------------------------------

Private Type BBox
    Valid As Boolean
    MinX As Double
    MinY As Double
    MaxX As Double
    MaxY As Double
End Type

Private Type PageInfo
    Valid As Boolean
    Source As String
    Paper As String
    IsSheet As Boolean
    WidthMM As Double
    HeightMM As Double
    MeasuredWidthMM As Double
    MeasuredHeightMM As Double
    Box As BBox
End Type

Private m_fso As Object
Private m_logPath As String

'==============================================================================
' Publiczne punkty wejscia
'==============================================================================

' Glowne wejscie. Drukuje wszystkie modele z aktywnego DGN do osobnych PDF-ow.
Public Sub Main()
    PrintCorrectSize
End Sub

' Pokazuje, co zostaloby wydrukowane, bez tworzenia PDF-ow.
Public Sub Preview()
    PreviewCorrectSize
End Sub

' Testuje logike rozmiarow i nazw plikow bez drukowania i bez DGN.
' Uruchom: vba run [PrintCorrectSize]PrintCorrectSize.SelfTest
Public Sub SelfTest()
    Dim failures As Long
    Dim total As Long
    Dim result As String
    Dim polishName As String
    Dim failNames As String

    failures = 0
    total = 0
    failNames = ""

    CheckSnap "A4 212x299", 212#, 299#, True, "A4", 210#, 297#, failures, total, failNames
    CheckSnap "A4 poziomo 299x212", 299#, 212#, True, "A4", 297#, 210#, failures, total, failNames
    CheckSnap "A3 300x423", 300#, 423#, True, "A3", 297#, 420#, failures, total, failNames
    CheckSnap "A4 za daleko 216x300", 216#, 300#, False, "", 0#, 0#, failures, total, failNames
    CheckSnap "A2 bez zaokraglania 420x594", 420#, 594#, False, "", 0#, 0#, failures, total, failNames
    CheckSnap "B1 bez zaokraglania 707x1000", 707#, 1000#, False, "", 0#, 0#, failures, total, failNames

    total = total + 1
    If KeyinNumber(210#) <> "210.000" Then
        failures = failures + 1
        failNames = failNames & "KeyinNumber; "
    End If

    total = total + 1
    If Not MatchesAny("M_RAMKA", "RAMKA") Then
        failures = failures + 1
        failNames = failNames & "MatchesAny M_RAMKA; "
    End If
    total = total + 1
    If Not MatchesAny(" _RAMKA", "RAMKA") Then
        failures = failures + 1
        failNames = failNames & "MatchesAny _RAMKA; "
    End If

    total = total + 1
    If SanitizeFileName("A/B:C") <> "A_B_C" Then
        failures = failures + 1
        failNames = failNames & "Sanitize A/B:C; "
    End If
    total = total + 1
    If SanitizeFileName("   ") <> "model" Then
        failures = failures + 1
        failNames = failNames & "Sanitize spaces; "
    End If
    total = total + 1
    If SanitizeFileName("CON") <> "_CON" Then
        failures = failures + 1
        failNames = failNames & "Sanitize CON; "
    End If

    polishName = "Mapa " & ChrW(&H141) & ChrW(&HF3) & "d" & ChrW(&H17A) & " 1"
    total = total + 1
    If SanitizeFileName(polishName) <> "Mapa_Lodz_1" Then
        failures = failures + 1
        failNames = failNames & "Sanitize Polish; "
    End If

    If failures = 0 Then
        result = "SelfTest OK." & vbCrLf & vbCrLf & _
                 "Sprawdzone testy: " & total & "/" & total
    Else
        result = "SelfTest: BLAD" & vbCrLf & vbCrLf & _
                 "Nieudane testy: " & failures & " z " & total & vbCrLf & vbCrLf & _
                 failNames
    End If

    MsgBox result, IIf(failures = 0, vbInformation, vbExclamation), "PrintCorrectSize - SelfTest"
End Sub

Private Sub CheckSnap(ByVal testName As String, ByVal wMM As Double, ByVal hMM As Double, _
                      ByVal expectSnap As Boolean, ByVal expectPaper As String, _
                      ByVal expectW As Double, ByVal expectH As Double, _
                      ByRef failures As Long, ByRef total As Long, ByRef failNames As String)
    Dim outW As Double
    Dim outH As Double
    Dim paperName As String
    Dim gotSnap As Boolean

    total = total + 1
    gotSnap = SnapToStandard(wMM, hMM, outW, outH, paperName)

    If gotSnap <> expectSnap Then
        failures = failures + 1
        failNames = failNames & testName & " (snap); "
        Exit Sub
    End If

    If expectSnap Then
        If paperName <> expectPaper Then
            failures = failures + 1
            failNames = failNames & testName & " (paper); "
        ElseIf Abs(outW - expectW) > 0.01 Or Abs(outH - expectH) > 0.01 Then
            failures = failures + 1
            failNames = failNames & testName & " (size); "
        End If
    End If
End Sub

' Drukuje wszystkie modele z aktywnego DGN do osobnych PDF-ow.
Public Sub PrintCorrectSize()
    Dim df As DesignFile
    Dim models As ModelReferences
    Dim originalModel As ModelReference
    Dim printModels() As ModelReference
    Dim infos() As PageInfo
    Dim pi As PageInfo
    Dim names() As String
    Dim paths() As String
    Dim m As ModelReference
    Dim i As Long
    Dim n As Long
    Dim outFolder As String
    Dim driverArg As String
    Dim summary As String
    Dim result As String
    Dim okCount As Long
    Dim failCount As Long
    Dim skipped As Long
    Dim answer As VbMsgBoxResult

    If Not Application.HasActiveDesignFile Then
        MsgBox "Najpierw otworz plik DGN.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set df = ActiveDesignFile
    Set models = df.Models
    If models.Count = 0 Then
        MsgBox "Aktywny plik DGN nie ma modeli.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    If Len(df.FullName) = 0 Then
        MsgBox "Zapisz plik DGN przed drukowaniem.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    On Error Resume Next
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    Err.Clear
    If m_fso Is Nothing Then
        MsgBox "Nie mozna uruchomic Scripting.FileSystemObject.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    If CHECK_PDF_DRIVER Then
        If Not PrintDriverExists(PDF_DRIVER) Then
            MsgBox "Nie znaleziono sterownika Bentley PDF '" & PDF_DRIVER & "'." & vbCrLf & vbCrLf & _
                   "Sprawdz, czy sterownik jest zainstalowany i czy MS_PLTCFG_PATH " & _
                   "wskazuje jego folder. Mozesz zmienic PDF_DRIVER na poczatku makra " & _
                   "albo ustawic CHECK_PDF_DRIVER na False.", _
                   vbExclamation, "PrintCorrectSize"
            Exit Sub
        End If
    End If

    outFolder = GetOutputFolder(df)
    If outFolder = "" Then Exit Sub

    If Not EnsureFolder(outFolder) Then
        MsgBox "Nie mozna utworzyc folderu wyjsciowego:" & vbCrLf & outFolder, vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    m_logPath = ""
    If WRITE_LOG Then
        m_logPath = m_fso.BuildPath(outFolder, _
                    "PrintCorrectSize_" & SanitizeFileName(m_fso.GetBaseName(df.FullName)) & ".log")
        On Error Resume Next
        m_fso.CreateTextFile m_logPath, True, True
        On Error GoTo 0
        Err.Clear
    End If

    Set originalModel = ActiveModelReference
    BuildPrintList models, printModels, n

    If n = 0 Then
        MsgBox "Nie znaleziono modeli do wydruku.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    ReDim infos(0 To n - 1)
    ReDim names(0 To n - 1)
    ReDim paths(0 To n - 1)

    ' Skanuj modele. Aktywacja jest potrzebna dla zaznaczenia i dla pewnego
    ' skanowania kazdego modelu w V8i.
    For i = 1 To n
        Set m = printModels(i - 1)
        m.Activate
        names(i - 1) = m.Name
        GetPageInfo m, pi
        infos(i - 1) = pi
        paths(i - 1) = MakePdfPath(df, m, n, outFolder)
    Next i
    originalModel.Activate

    ' Zbuduj krotki raport i popros o potwierdzenie.
    summary = "Aktywny DGN: " & df.FullName & vbCrLf & _
              "Folder wyjsciowy: " & outFolder & vbCrLf & vbCrLf

    For i = 0 To n - 1
        If infos(i).Valid Then
            summary = summary & Format$(i + 1, "00") & ". " & names(i) & _
                      "  ->  " & Format$(infos(i).WidthMM, "0.0") & " x " & _
                      Format$(infos(i).HeightMM, "0.0") & " mm  [" & infos(i).Source & "]"
            If infos(i).Paper <> "WLASNY" Then
                summary = summary & "  ->  " & infos(i).Paper & _
                          " (zmierzono " & Format$(infos(i).MeasuredWidthMM, "0.0") & " x " & _
                          Format$(infos(i).MeasuredHeightMM, "0.0") & " mm)"
            End If
            summary = summary & vbCrLf
        Else
            summary = summary & Format$(i + 1, "00") & ". " & names(i) & _
                      "  ->  POMINIETY (brak geometrii)" & vbCrLf
        End If
    Next i

    summary = summary & vbCrLf & "Utworzyc te pliki PDF?"
    answer = MsgBox(summary, vbOKCancel + vbInformation, "PrintCorrectSize")
    If answer <> vbOK Then Exit Sub

    ' Jednorazowa konfiguracja drukowania.
    driverArg = PDF_DRIVER
    If InStr(driverArg, " ") > 0 Then driverArg = Chr$(34) & driverArg & Chr$(34)
    CadInputQueue.SendKeyin "PRINT DRIVER " & driverArg
    CadInputQueue.SendKeyin "PRINT DESTINATION PLOTFILE"
    CadInputQueue.SendKeyin "PRINT COLORMODE COLOR"
    CadInputQueue.SendKeyin "PRINT AUTOCENTER ON"
    ' Nie dodawaj wlasnej ramki wydruku MicroStation ani obrysu fence'a.
    CadInputQueue.SendKeyin "PRINT ATTRIBUTES PLOTBORDER OFF"
    CadInputQueue.SendKeyin "PRINT ATTRIBUTES FENCEBOUNDARY OFF"

    LogLine "==== " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===="
    LogLine "DGN: " & df.FullName
    LogLine "Sterownik: " & PDF_DRIVER

    For i = 0 To n - 1
        If infos(i).Valid Then
            Set m = printModels(i)
            m.Activate

            If PrintOneModel(m, infos(i), paths(i)) Then
                okCount = okCount + 1
                LogLine "OK   " & names(i) & " | " & _
                        Format$(infos(i).WidthMM, "0.0") & "x" & Format$(infos(i).HeightMM, "0.0") & _
                        " mm | " & infos(i).Paper & " | zmierzono " & _
                        Format$(infos(i).MeasuredWidthMM, "0.0") & "x" & _
                        Format$(infos(i).MeasuredHeightMM, "0.0") & " mm | " & _
                        infos(i).Source & " | " & paths(i)
            Else
                failCount = failCount + 1
                LogLine "BLAD " & names(i) & " | " & paths(i)
            End If
        Else
            skipped = skipped + 1
            LogLine "POMINIETY " & names(i)
        End If
    Next i

    ' Zostaw plik w takim stanie, w jakim go zastalismy.
    On Error Resume Next
    ActiveDesignFile.Fence.Undefine
    On Error GoTo 0
    originalModel.Activate

    result = "Gotowe." & vbCrLf & vbCrLf & _
             "Utworzono: " & okCount & vbCrLf & _
             "Bledy:     " & failCount & vbCrLf & _
             "Pominiete: " & skipped & vbCrLf & vbCrLf & _
             "Folder: " & outFolder

    If failCount > 0 Then
        result = result & vbCrLf & vbCrLf & _
                 "Sprawdz log PrintCorrectSize i upewnij sie, ze sterownik Bentley PDF " & _
                 "(" & PDF_DRIVER & ") jest zainstalowany i pozwala edytowac rozmiar papieru."
    End If

    MsgBox result, IIf(failCount > 0, vbExclamation, vbInformation), "PrintCorrectSize"
End Sub

' Pokazuje wykryte rozmiary stron bez drukowania.
Public Sub PreviewCorrectSize()
    Dim df As DesignFile
    Dim models As ModelReferences
    Dim printModels() As ModelReference
    Dim m As ModelReference
    Dim originalModel As ModelReference
    Dim pi As PageInfo
    Dim i As Long
    Dim n As Long
    Dim report As String

    If Not Application.HasActiveDesignFile Then
        MsgBox "Najpierw otworz plik DGN.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set df = ActiveDesignFile
    Set models = df.Models
    If models.Count = 0 Then
        MsgBox "Aktywny plik DGN nie ma modeli.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set originalModel = ActiveModelReference
    BuildPrintList models, printModels, n

    If n = 0 Then
        MsgBox "Nie znaleziono modeli do wydruku.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    report = "Aktywny DGN: " & df.FullName & vbCrLf & vbCrLf

    For i = 1 To n
        Set m = printModels(i - 1)
        m.Activate
        GetPageInfo m, pi

        report = report & Format$(i, "00") & ". " & m.Name & vbCrLf
        If pi.Valid Then
            report = report & "     " & Format$(pi.WidthMM, "0.0") & " x " & _
                     Format$(pi.HeightMM, "0.0") & " mm  [" & pi.Source & "]"
            If pi.Paper <> "WLASNY" Then
                report = report & "  ->  " & pi.Paper & _
                         " (zmierzono " & Format$(pi.MeasuredWidthMM, "0.0") & " x " & _
                         Format$(pi.MeasuredHeightMM, "0.0") & " mm)"
            End If
            report = report & vbCrLf
        Else
            report = report & "     POMINIETY - brak geometrii" & vbCrLf
        End If
    Next i
    originalModel.Activate

    MsgBox report, vbInformation, "PrintCorrectSize - podglad"
End Sub

'==============================================================================
' Wybor modeli
'==============================================================================

Private Sub BuildPrintList(ByVal models As ModelReferences, _
                           ByRef result() As ModelReference, ByRef count As Long)
    Dim hasSheet As Boolean
    Dim i As Long

    hasSheet = False
    For i = 1 To models.Count
        If models(i).Type = msdModelTypeSheet Then
            hasSheet = True
            Exit For
        End If
    Next i

    ReDim result(0 To models.Count - 1)
    count = 0

    For i = 1 To models.Count
        If ShouldPrintModel(models(i), hasSheet) Then
            Set result(count) = models(i)
            count = count + 1
        End If
    Next i

    If count = 0 Then ReDim result(0 To 0)
End Sub

Private Function ShouldPrintModel(ByVal m As ModelReference, ByVal hasSheet As Boolean) As Boolean
    If hasSheet And PRINT_ONLY_SHEETS_IF_PRESENT Then
        ShouldPrintModel = (m.Type = msdModelTypeSheet)
    Else
        ShouldPrintModel = True
    End If
End Function

'==============================================================================
' Wykrywanie rozmiaru strony
'==============================================================================

Private Sub GetPageInfo(ByVal m As ModelReference, ByRef pi As PageInfo)
    Dim sd As SheetDefinition
    Dim b As BBox

    pi.Valid = False
    pi.Source = ""
    pi.Paper = ""
    pi.IsSheet = (m.Type = msdModelTypeSheet)
    pi.Box.Valid = False

    ' 1. Model arkusza (sheet model): definicja arkusza jest nadrzednym
    '    rozmiarem strony.
    If pi.IsSheet Then
        On Error Resume Next
        Err.Clear
        Set sd = m.GetSheetDefinition
        If Err.Number = 0 Then
            If Not sd Is Nothing Then
                If sd.Width > 0# And sd.Height > 0# Then
                    SetPageSize pi, UORsToMM(m, sd.Width), UORsToMM(m, sd.Height), "ARKUSZ"
                End If
            End If
        End If
        On Error GoTo 0
        Err.Clear
        If pi.Valid Then Exit Sub

        ' Model arkusza bez poprawnej definicji trzeba drukowac z fence'a,
        ' tak jak zwykly model rysunkowy.
        pi.IsSheet = False
    End If

    ' 2. Elementy na levelu ramki. To normalna sciezka dla modeli
    '    rysunkowych z ramka na M_RAMKA albo _RAMKA.
    If Len(Trim$(BORDER_LEVEL_HINTS)) > 0 Then
        On Error Resume Next
        Err.Clear
        b = BBoxFromBorderLevels(m)
        If Err.Number = 0 And b.Valid Then
            SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                        UORsToMM(m, b.MaxY - b.MinY), "LEVEL RAMKI"
            If pi.Valid Then
                pi.Box = b
                Exit Sub
            End If
        End If
        On Error GoTo 0
        Err.Clear
    End If

    ' 3. Reczne wymuszenie: zaznaczone elementy.
    If USE_SELECTION_AS_PAGE Then
        On Error Resume Next
        Err.Clear
        If m.AnyElementsSelected Then
            b = BBoxFromEnumerator(m.GetSelectedElements)
            If b.Valid Then
                SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                            UORsToMM(m, b.MaxY - b.MinY), "ZAZNACZENIE"
                If pi.Valid Then
                    pi.Box = b
                    Exit Sub
                End If
            End If
        End If
        On Error GoTo 0
        Err.Clear
    End If

    ' 4. Najwiekszy zamkniety element (zwykle obrys arkusza).
    On Error Resume Next
    Err.Clear
    b = BBoxFromLargestClosedElement(m)
    If Err.Number = 0 And b.Valid Then
        SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                    UORsToMM(m, b.MaxY - b.MinY), "ZAMKNIETY ELEMENT"
        If pi.Valid Then
            pi.Box = b
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Err.Clear

    ' 5. Rozwiazanie awaryjne: wszystkie elementy graficzne.
    On Error Resume Next
    Err.Clear
    b = BBoxFromAllElements(m)
    If Err.Number = 0 And b.Valid Then
        SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                    UORsToMM(m, b.MaxY - b.MinY), "GRANICE"
        If pi.Valid Then
            pi.Box = b
        End If
    End If
    On Error GoTo 0
    Err.Clear
End Sub

Private Function BBoxFromEnumerator(ByVal ee As ElementEnumerator) As BBox
    Dim b As BBox
    Dim el As Element
    Dim r As Range3d

    Do While ee.MoveNext
        Set el = ee.Current
        If el.IsGraphical Then
            r = el.Range(False)
            BBoxAdd b, r
        End If
    Loop

    BBoxFromEnumerator = b
End Function

Private Function BBoxFromBorderLevels(ByVal m As ModelReference) As BBox
    Dim allOnLevels As BBox
    Dim closedBest As BBox
    Dim bestArea As Double
    Dim area As Double
    Dim ee As ElementEnumerator
    Dim el As Element
    Dim r As Range3d
    Dim levelName As String

    bestArea = 0#
    Set ee = m.Scan(Nothing)
    Do While ee.MoveNext
        Set el = ee.Current
        If el.IsGraphical Then
            levelName = ""
            On Error Resume Next
            Err.Clear
            levelName = el.Level.Name
            On Error GoTo 0
            Err.Clear

            If MatchesAny(levelName, BORDER_LEVEL_HINTS) Then
                r = el.Range(False)

                ' Tekst i wymiary na levelu ramki nie moga powiekszac strony.
                ' To zwykle notatki umieszczone poza ramka.
                If Not el.IsTextElement And Not el.IsTextNodeElement And _
                   Not el.IsDimensionElement Then
                    BBoxAdd allOnLevels, r
                End If

                ' Preferuj najwiekszy zamkniety element (ksztalt ramki) nad
                ' suma wszystkiego na levelu. To ignoruje notatki i inne
                ' elementy, ktore przypadkowo sa na levelu ramki.
                If el.IsClosedElement Then
                    area = (r.High.X - r.Low.X) * (r.High.Y - r.Low.Y)
                    If area > bestArea Then
                        bestArea = area
                        BBoxInit closedBest
                        BBoxAdd closedBest, r
                    End If
                End If
            End If
        End If
    Loop

    ' Wybierz wiekszy kandydat: ksztalt ramki albo suma elementow innych niz
    ' tekst na levelu ramki. Zamknieta ramka zwykle wygrywa; gdy ramka jest
    ' z linii, suma wygrywa z mniejszym zamknietym blokiem tytulowym.
    If closedBest.Valid And allOnLevels.Valid Then
        If BBoxArea(closedBest) > BBoxArea(allOnLevels) Then
            BBoxFromBorderLevels = closedBest
        Else
            BBoxFromBorderLevels = allOnLevels
        End If
    ElseIf closedBest.Valid Then
        BBoxFromBorderLevels = closedBest
    Else
        BBoxFromBorderLevels = allOnLevels
    End If
End Function

Private Function BBoxArea(ByRef b As BBox) As Double
    If Not b.Valid Then
        BBoxArea = 0#
    Else
        BBoxArea = (b.MaxX - b.MinX) * (b.MaxY - b.MinY)
    End If
End Function

Private Function BBoxFromLargestClosedElement(ByVal m As ModelReference) As BBox
    Dim b As BBox
    Dim best As BBox
    Dim bestArea As Double
    Dim ee As ElementEnumerator
    Dim el As Element
    Dim r As Range3d
    Dim area As Double

    bestArea = 0#
    Set ee = m.Scan(Nothing)
    Do While ee.MoveNext
        Set el = ee.Current
        If el.IsGraphical Then
            If el.IsClosedElement Then
                r = el.Range(False)
                area = (r.High.X - r.Low.X) * (r.High.Y - r.Low.Y)
                If area > bestArea Then
                    bestArea = area
                    BBoxInit best
                    BBoxAdd best, r
                End If
            End If
        End If
    Loop

    BBoxFromLargestClosedElement = best
End Function

Private Function BBoxFromAllElements(ByVal m As ModelReference) As BBox
    Dim b As BBox
    Dim ee As ElementEnumerator
    Dim el As Element
    Dim r As Range3d

    Set ee = m.Scan(Nothing)
    Do While ee.MoveNext
        Set el = ee.Current
        If el.IsGraphical Then
            r = el.Range(False)
            BBoxAdd b, r
        End If
    Loop

    BBoxFromAllElements = b
End Function

Private Sub BBoxInit(ByRef b As BBox)
    b.Valid = False
    b.MinX = 0#
    b.MinY = 0#
    b.MaxX = 0#
    b.MaxY = 0#
End Sub

Private Sub BBoxAdd(ByRef b As BBox, ByRef r As Range3d)
    ' Ignoruj puste lub niepoprawne zakresy.
    If r.High.X < r.Low.X Then Exit Sub
    If r.High.Y < r.Low.Y Then Exit Sub

    If Not b.Valid Then
        b.MinX = r.Low.X
        b.MinY = r.Low.Y
        b.MaxX = r.High.X
        b.MaxY = r.High.Y
        b.Valid = True
    Else
        If r.Low.X < b.MinX Then b.MinX = r.Low.X
        If r.Low.Y < b.MinY Then b.MinY = r.Low.Y
        If r.High.X > b.MaxX Then b.MaxX = r.High.X
        If r.High.Y > b.MaxY Then b.MaxY = r.High.Y
    End If
End Sub

Private Function IsSanePage(ByVal wMM As Double, ByVal hMM As Double) As Boolean
    IsSanePage = False
    If wMM < MIN_PAGE_MM Or hMM < MIN_PAGE_MM Then Exit Function
    If wMM > MAX_PAGE_MM Or hMM > MAX_PAGE_MM Then Exit Function
    IsSanePage = True
End Function

Private Sub SetPageSize(ByRef pi As PageInfo, ByVal wMM As Double, ByVal hMM As Double, _
                        ByVal source As String)
    Dim stdW As Double
    Dim stdH As Double
    Dim paperName As String

    pi.MeasuredWidthMM = wMM
    pi.MeasuredHeightMM = hMM
    pi.WidthMM = wMM
    pi.HeightMM = hMM
    pi.Source = source
    pi.Paper = "WLASNY"

    If SnapToStandard(wMM, hMM, stdW, stdH, paperName) Then
        pi.WidthMM = stdW
        pi.HeightMM = stdH
        pi.Paper = paperName
    End If

    pi.Valid = IsSanePage(pi.WidthMM, pi.HeightMM)
End Sub

' Zwraca True, gdy zmierzona strona miesci sie w tolerancji jednego z
' rozmiarow z STANDARD_SIZES_TO_SNAP. Zwrocona szerokosc/wysokosc zachowuje
' orientacje zmierzonej strony.
Private Function SnapToStandard(ByVal wMM As Double, ByVal hMM As Double, _
                                ByRef outW As Double, ByRef outH As Double, _
                                ByRef paperName As String) As Boolean
    Dim names() As String
    Dim i As Long
    Dim stdW As Double
    Dim stdH As Double
    Dim measuredShort As Double
    Dim measuredLong As Double
    Dim standardShort As Double
    Dim standardLong As Double
    Dim candidate As String

    SnapToStandard = False
    paperName = "WLASNY"

    If Len(Trim$(STANDARD_SIZES_TO_SNAP)) = 0 Then Exit Function

    measuredShort = wMM
    measuredLong = hMM
    If measuredShort > measuredLong Then
        measuredShort = hMM
        measuredLong = wMM
    End If

    names = Split(STANDARD_SIZES_TO_SNAP, ",")
    For i = LBound(names) To UBound(names)
        candidate = Trim$(CStr(names(i)))
        If Len(candidate) > 0 Then
            If PaperDimensions(candidate, stdW, stdH) Then
                standardShort = stdW
                standardLong = stdH
                If standardShort > standardLong Then
                    standardShort = stdH
                    standardLong = stdW
                End If

                If Abs(measuredShort - standardShort) <= STANDARD_SIZE_TOLERANCE_MM And _
                   Abs(measuredLong - standardLong) <= STANDARD_SIZE_TOLERANCE_MM Then

                    If wMM > hMM Then
                        outW = standardLong
                        outH = standardShort
                    Else
                        outW = standardShort
                        outH = standardLong
                    End If

                    paperName = UCase$(candidate)
                    SnapToStandard = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' Rozmiary ISO A i B w milimetrach, orientacja pionowa.
Private Function PaperDimensions(ByVal paperName As String, _
                                 ByRef wMM As Double, ByRef hMM As Double) As Boolean
    PaperDimensions = True

    Select Case UCase$(Trim$(paperName))
        Case "A0"
            wMM = 841#
            hMM = 1189#
        Case "A1"
            wMM = 594#
            hMM = 841#
        Case "A2"
            wMM = 420#
            hMM = 594#
        Case "A3"
            wMM = 297#
            hMM = 420#
        Case "A4"
            wMM = 210#
            hMM = 297#
        Case "A5"
            wMM = 148#
            hMM = 210#
        Case "A6"
            wMM = 105#
            hMM = 148#
        Case "B0"
            wMM = 1000#
            hMM = 1414#
        Case "B1"
            wMM = 707#
            hMM = 1000#
        Case "B2"
            wMM = 500#
            hMM = 707#
        Case "B3"
            wMM = 353#
            hMM = 500#
        Case "B4"
            wMM = 250#
            hMM = 353#
        Case "B5"
            wMM = 176#
            hMM = 250#
        Case "B6"
            wMM = 125#
            hMM = 176#
        Case Else
            PaperDimensions = False
    End Select
End Function

Private Function UORsToMM(ByVal m As ModelReference, ByVal uors As Double) As Double
    Dim mu As MeasurementUnit
    Dim masterUnits As Double
    Dim metres As Double

    If m.UORsPerMasterUnit = 0# Then
        UORsToMM = 0#
        Exit Function
    End If

    mu = m.MasterUnit
    masterUnits = uors / m.UORsPerMasterUnit

    If mu.UnitsPerBaseNumerator = 0# Then
        UORsToMM = 0#
        Exit Function
    End If

    metres = masterUnits * mu.UnitsPerBaseDenominator / mu.UnitsPerBaseNumerator
    UORsToMM = metres * 1000#
End Function

'==============================================================================
' Drukowanie
'==============================================================================

Private Function PrintOneModel(ByVal m As ModelReference, ByRef pi As PageInfo, _
                               ByVal pdfPath As String) As Boolean
    Dim shortMM As Double
    Dim longMM As Double
    Dim orientation As String
    Dim keyinPath As String
    Dim pageBox As BBox

    PrintOneModel = False
    On Error GoTo errHandler

    ' Najpierw wybierz obszar druku. Dla modelu arkusza powoduje to rowniez
    ' uzycie obrysu arkusza przez MicroStation.
    If pi.IsSheet Then
        CadInputQueue.SendKeyin "PRINT BOUNDARY SHEET"
    Else
        pageBox = pi.Box
        If Not DefineFenceForPage(m, pageBox) Then
            LogLine "  nie mozna zdefiniowac fence'a dla " & m.Name
            Exit Function
        End If
        CadInputQueue.SendKeyin "PRINT BOUNDARY FENCE"
    End If

    ' Sterownik Bentley uzywa formularza pionowego plus orientacji. Ustaw
    ' krotszy bok jako X, dluzszy jako Y, a potem obroc, jesli trzeba.
    shortMM = pi.WidthMM
    longMM = pi.HeightMM
    If shortMM > longMM Then
        shortMM = pi.HeightMM
        longMM = pi.WidthMM
    End If

    If pi.WidthMM > pi.HeightMM Then
        orientation = "LANDSCAPE"
    Else
        orientation = "PORTRAIT"
    End If

    CadInputQueue.SendKeyin "PRINT UNITS MM"
    ' Ustaw formularz w orientacji pionowej, a potem obroc strone.
    CadInputQueue.SendKeyin "PRINT ORIENTATION PORTRAIT"
    CadInputQueue.SendKeyin "PRINT XFORM " & KeyinNumber(shortMM) & " MM"
    CadInputQueue.SendKeyin "PRINT YFORM " & KeyinNumber(longMM) & " MM"
    CadInputQueue.SendKeyin "PRINT ORIENTATION " & orientation
    CadInputQueue.SendKeyin "PRINT AUTOCENTER ON"
    CadInputQueue.SendKeyin "PRINT MAXIMIZE"

    ' Usun stary plik, zeby przekroczenie czasu nie zostalo uznane za sukces.
    On Error Resume Next
    If m_fso.FileExists(pdfPath) Then m_fso.DeleteFile pdfPath, True
    On Error GoTo errHandler
    Err.Clear
    If m_fso.FileExists(pdfPath) Then
        LogLine "  nie mozna usunac starego PDF-a: " & pdfPath
        Exit Function
    End If

    ' Uzyj krotkiej sciezki folderu i nazwy pliku bez spacji. To unika
    ' niejednoznacznosci, gdy parser key-inow widzi sciezke ze spacjami.
    keyinPath = MakeKeyinPath(pdfPath)
    CadInputQueue.SendKeyin "PRINT EXECUTE " & keyinPath

    PrintOneModel = WaitForPdf(pdfPath, PDF_TIMEOUT_SEC)
    Exit Function

errHandler:
    LogLine "  blad " & Err.Number & ": " & Err.Description
    PrintOneModel = False
End Function

Private Function DefineFenceForPage(ByVal m As ModelReference, ByRef b As BBox) As Boolean
    Dim fence As Fence
    Dim pts(0 To 3) As Point3d
    Dim scaleToMaster As Double

    DefineFenceForPage = False
    If Not b.Valid Then Exit Function
    If m.UORsPerMasterUnit = 0# Then Exit Function

    scaleToMaster = 1# / m.UORsPerMasterUnit

    pts(0).X = b.MinX * scaleToMaster
    pts(0).Y = b.MinY * scaleToMaster
    pts(0).Z = 0#

    pts(1).X = b.MaxX * scaleToMaster
    pts(1).Y = b.MinY * scaleToMaster
    pts(1).Z = 0#

    pts(2).X = b.MaxX * scaleToMaster
    pts(2).Y = b.MaxY * scaleToMaster
    pts(2).Z = 0#

    pts(3).X = b.MinX * scaleToMaster
    pts(3).Y = b.MaxY * scaleToMaster
    pts(3).Z = 0#

    On Error Resume Next
    Set fence = ActiveDesignFile.Fence
    fence.Undefine
    Err.Clear
    fence.DefineFromModelPoints 1, pts
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    DefineFenceForPage = True
End Function

Private Function WaitForPdf(ByVal path As String, ByVal timeoutSec As Long) As Boolean
    Dim started As Double
    Dim lastSize As Double
    Dim stableCount As Long
    Dim sizeNow As Double

    WaitForPdf = False
    started = Timer
    lastSize = -1#
    stableCount = 0

    Do
        DoEvents

        If m_fso.FileExists(path) Then
            On Error Resume Next
            Err.Clear
            sizeNow = m_fso.GetFile(path).Size
            On Error GoTo 0
            Err.Clear

            If sizeNow > 0# Then
                If sizeNow = lastSize Then
                    stableCount = stableCount + 1
                Else
                    stableCount = 0
                    lastSize = sizeNow
                End If

                ' Rozmiar pliku bez zmian przez ~0,6 s -> zakladamy, ze PDF jest gotowy.
                If stableCount >= 3 Then
                    WaitForPdf = True
                    Exit Function
                End If
            End If
        End If

        If ElapsedSeconds(started) > timeoutSec Then Exit Function
        Pause 0.2
    Loop
End Function

Private Sub Pause(ByVal seconds As Double)
    Dim started As Double
    started = Timer
    Do While ElapsedSeconds(started) < seconds
        DoEvents
    Loop
End Sub

Private Function ElapsedSeconds(ByVal started As Double) As Double
    Dim nowValue As Double
    nowValue = Timer
    If nowValue >= started Then
        ElapsedSeconds = nowValue - started
    Else
        ' Przejscie przez polnoc.
        ElapsedSeconds = (86400# - started) + nowValue
    End If
End Function

'==============================================================================
' Pliki, nazwy i drobne funkcje pomocnicze
'==============================================================================

Private Function PrintDriverExists(ByVal driverName As String) As Boolean
    Dim paths As String
    Dim parts As Variant
    Dim i As Long
    Dim folder As String

    PrintDriverExists = False
    If Len(Trim$(driverName)) = 0 Then Exit Function

    ' Pozwol, zeby PDF_DRIVER byl pelna sciezka albo sama nazwa pliku.
    If m_fso.FileExists(driverName) Then
        PrintDriverExists = True
        Exit Function
    End If

    On Error Resume Next
    If Not ActiveWorkspace.IsConfigurationVariableDefined("MS_PLTCFG_PATH") Then
        Err.Clear
        Exit Function
    End If
    paths = ActiveWorkspace.ConfigurationVariableValue("MS_PLTCFG_PATH")
    On Error GoTo 0
    Err.Clear

    parts = Split(paths, ";")
    For i = LBound(parts) To UBound(parts)
        folder = Trim$(CStr(parts(i)))
        If Len(folder) > 0 Then
            If m_fso.FileExists(m_fso.BuildPath(folder, driverName)) Then
                PrintDriverExists = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetOutputFolder(ByVal df As DesignFile) As String
    Dim dgnFolder As String
    Dim outFolder As String

    dgnFolder = m_fso.GetParentFolderName(df.FullName)
    If Len(OUTPUT_SUBFOLDER) = 0 Then
        outFolder = dgnFolder
    Else
        outFolder = m_fso.BuildPath(dgnFolder, OUTPUT_SUBFOLDER)
    End If

    GetOutputFolder = outFolder
End Function

Private Function EnsureFolder(ByVal folder As String) As Boolean
    EnsureFolder = False
    On Error Resume Next
    If m_fso.FolderExists(folder) Then
        EnsureFolder = True
    Else
        m_fso.CreateFolder folder
        EnsureFolder = m_fso.FolderExists(folder)
    End If
    On Error GoTo 0
End Function

Private Function MakePdfPath(ByVal df As DesignFile, ByVal m As ModelReference, _
                             ByVal modelCount As Long, ByVal outFolder As String) As String
    Dim baseName As String
    Dim fileName As String

    baseName = SanitizeFileName(m_fso.GetBaseName(df.FullName))

    If modelCount = 1 Then
        fileName = baseName & ".pdf"
    Else
        fileName = baseName & "_" & SanitizeFileName(m.Name) & ".pdf"
    End If

    MakePdfPath = m_fso.BuildPath(outFolder, fileName)
End Function

' Zwraca sciezke odpowiednia dla parsera key-inow. Krotka sciezka folderu
' usuwa spacje z nazw katalogow; nazwa pliku jest juz oczyszczona przez
' SanitizeFileName. Jesli gdzies zostanie spacja, sciezka jest ujeta w cudzyslow.
Private Function MakeKeyinPath(ByVal path As String) As String
    Dim folder As String
    Dim fileName As String
    Dim shortFolder As String

    folder = m_fso.GetParentFolderName(path)
    fileName = m_fso.GetFileName(path)

    On Error Resume Next
    shortFolder = m_fso.GetFolder(folder).ShortPath
    On Error GoTo 0
    Err.Clear
    If Len(shortFolder) = 0 Then shortFolder = folder

    MakeKeyinPath = m_fso.BuildPath(shortFolder, fileName)

    If InStr(MakeKeyinPath, " ") > 0 Then
        MakeKeyinPath = Chr$(34) & MakeKeyinPath & Chr$(34)
    End If
End Function

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant
    Dim i As Long
    Dim result As String

    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", " ", ";")
    result = Trim$(s)
    If Len(result) = 0 Then
        SanitizeFileName = "model"
        Exit Function
    End If

    For i = LBound(badChars) To UBound(badChars)
        result = Replace(result, CStr(badChars(i)), "_")
    Next i

    ' Nazwy plikow dla key-inow zapisujemy bez polskich znakow. ChrW zachowuje
    ' plik zrodlowy w ASCII, a wynik jest bezpieczny dla parsera V8i.
    result = Replace(result, ChrW(&H105), "a")   ' a z ogonkiem
    result = Replace(result, ChrW(&H104), "A")
    result = Replace(result, ChrW(&H107), "c")   ' c z kreska
    result = Replace(result, ChrW(&H106), "C")
    result = Replace(result, ChrW(&H119), "e")   ' e z ogonkiem
    result = Replace(result, ChrW(&H118), "E")
    result = Replace(result, ChrW(&H142), "l")   ' l z kreska
    result = Replace(result, ChrW(&H141), "L")
    result = Replace(result, ChrW(&H144), "n")   ' n z kreska
    result = Replace(result, ChrW(&H143), "N")
    result = Replace(result, ChrW(&HF3), "o")    ' o z kreska
    result = Replace(result, ChrW(&HD3), "O")
    result = Replace(result, ChrW(&H15B), "s")   ' s z kreska
    result = Replace(result, ChrW(&H15A), "S")
    result = Replace(result, ChrW(&H17A), "z")   ' z z kreska
    result = Replace(result, ChrW(&H179), "Z")
    result = Replace(result, ChrW(&H17C), "z")   ' z z kropka
    result = Replace(result, ChrW(&H17B), "Z")

    result = Trim$(result)

    ' Windows nie akceptuje nazw zakonczonych kropka.
    Do While Len(result) > 0
        If Right$(result, 1) = "." Then
            result = Left$(result, Len(result) - 1)
        Else
            Exit Do
        End If
    Loop

    If Len(result) = 0 Then result = "model"
    If IsReservedFileName(result) Then result = "_" & result
    If Len(result) > 120 Then result = Left$(result, 120)
    SanitizeFileName = result
End Function

Private Function IsReservedFileName(ByVal name As String) As Boolean
    Dim reserved As Variant
    Dim baseName As String
    Dim i As Long

    reserved = Array("CON", "PRN", "AUX", "NUL", _
                     "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", _
                     "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9")

    baseName = UCase$(name)
    If InStr(baseName, ".") > 0 Then
        baseName = Left$(baseName, InStr(baseName, ".") - 1)
    End If

    For i = LBound(reserved) To UBound(reserved)
        If baseName = CStr(reserved(i)) Then
            IsReservedFileName = True
            Exit Function
        End If
    Next i

    IsReservedFileName = False
End Function

Private Function KeyinNumber(ByVal value As Double) As String
    Dim s As String
    s = Format$(value, "0.000")
    s = Replace(s, ",", ".")
    KeyinNumber = s
End Function

Private Function MatchesAny(ByVal text As String, ByVal commaList As String) As Boolean
    Dim parts As Variant
    Dim i As Long
    Dim token As String

    MatchesAny = False
    If Len(text) = 0 Then Exit Function

    parts = Split(commaList, ",")
    For i = LBound(parts) To UBound(parts)
        token = Trim$(CStr(parts(i)))
        If Len(token) > 0 Then
            If InStr(1, text, token, vbTextCompare) > 0 Then
                MatchesAny = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub LogLine(ByVal text As String)
    Dim ts As Object

    If Len(m_logPath) = 0 Then Exit Sub
    On Error Resume Next
    Set ts = m_fso.OpenTextFile(m_logPath, 8, True, -1)
    If Not ts Is Nothing Then
        ts.WriteLine text
        ts.Close
    End If
    On Error GoTo 0
End Sub
