Attribute VB_Name = "PrintCorrectSize"
'==============================================================================
' PrintCorrectSize
'
' MicroStation V8i VBA macro.
'
' Prints every printable model of the active DGN to a separate PDF and sets
' the PDF page size to the size of that model. This avoids the usual problem
' where one paper size is used for a whole batch: A4 sheets become small and
' larger sheets are clipped.
'
' Page size is detected in this order:
'   1. Sheet model definition (Model Properties > Sheet > Size)
'   2. Elements on a frame level (name contains one of BORDER_LEVEL_HINTS)
'   3. A selection set in the model (if anything is selected)  -> fallback override
'   4. Largest closed element in the model
'   5. Extents of all graphical elements in the model
'
' A measured page within STANDARD_SIZE_TOLERANCE_MM of A4 or A3 is snapped to
' the exact ISO size. Other sizes keep their measured width and height.
'
' The macro uses MicroStation's built-in Bentley PDF driver (pdf.pltcfg) and
' the documented PRINT ... key-ins, so no external PDF printer is required.
'
' Target: MicroStation V8i (SELECTseries). VBA 6/7.
'==============================================================================

Option Explicit

'------------------------------------------------------------------------------
' User-editable settings
'------------------------------------------------------------------------------

' Bentley PDF plot driver. The file is searched through MS_PLTCFG_PATH.
Private Const PDF_DRIVER As String = "pdf.pltcfg"

' Check that PDF_DRIVER exists before touching the drawing. Set to False if
' your driver is found through a path that MS_PLTCFG_PATH does not list.
Private Const CHECK_PDF_DRIVER As Boolean = True

' Output sub-folder created next to the DGN. Use "" to write next to the DGN.
Private Const OUTPUT_SUBFOLDER As String = "PDF"

' Level names (or parts of names) used for the sheet border. Case-insensitive.
' Separate names with a comma. Leave "" to skip this detection step.
' "RAMKA" matches both "M_RAMKA" and " _RAMKA".
Private Const BORDER_LEVEL_HINTS As String = "RAMKA"

' If True, a selection set is used as a fallback when no frame is found.
' A sheet model's sheet definition and a frame level always win.
Private Const USE_SELECTION_AS_PAGE As Boolean = True

' If a DGN contains sheet models, print only the sheet models. Design models
' in such a file are usually the source geometry, not printable sheets.
' Default is False because this macro is intended for design models that each
' contain their own frame.
Private Const PRINT_ONLY_SHEETS_IF_PRESENT As Boolean = False

' Standard sizes that are snapped to the exact ISO paper size.
' Only these names are snapped; every other size keeps its measured width and
' height. Default: A4 and A3 with a 5 mm tolerance.
Private Const STANDARD_SIZES_TO_SNAP As String = "A4,A3"
Private Const STANDARD_SIZE_TOLERANCE_MM As Double = 5#

' Sanity limits for a detected page (millimetres).
Private Const MIN_PAGE_MM As Double = 40#
Private Const MAX_PAGE_MM As Double = 10000#

' How long to wait for one PDF before giving up (seconds).
Private Const PDF_TIMEOUT_SEC As Long = 180

' Write a small log next to the PDFs.
Private Const WRITE_LOG As Boolean = True

'------------------------------------------------------------------------------
' Private types and module state
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
' Public entry points
'==============================================================================

' Main entry point. Prints all models of the active DGN to separate PDFs.
Public Sub Main()
    PrintCorrectSize
End Sub

' Shows what would be printed, without creating PDFs.
Public Sub Preview()
    PreviewCorrectSize
End Sub

' Prints all models of the active DGN to separate PDFs.
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
    Dim summary As String
    Dim result As String
    Dim okCount As Long
    Dim failCount As Long
    Dim skipped As Long
    Dim answer As VbMsgBoxResult

    If Not Application.HasActiveDesignFile Then
        MsgBox "Open a DGN file first.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set df = ActiveDesignFile
    Set models = df.Models
    If models.Count = 0 Then
        MsgBox "The active DGN has no models.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    If Len(df.FullName) = 0 Then
        MsgBox "Save the DGN file before printing.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    On Error Resume Next
    Set m_fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    Err.Clear
    If m_fso Is Nothing Then
        MsgBox "Cannot start Scripting.FileSystemObject.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    If CHECK_PDF_DRIVER Then
        If Not PrintDriverExists(PDF_DRIVER) Then
            MsgBox "Cannot find the Bentley PDF driver '" & PDF_DRIVER & "'." & vbCrLf & vbCrLf & _
                   "Check that the driver is installed and that MS_PLTCFG_PATH points " & _
                   "to its folder. You can change PDF_DRIVER at the top of the macro, " & _
                   "or set CHECK_PDF_DRIVER to False.", _
                   vbExclamation, "PrintCorrectSize"
            Exit Sub
        End If
    End If

    outFolder = GetOutputFolder(df)
    If outFolder = "" Then Exit Sub

    If Not EnsureFolder(outFolder) Then
        MsgBox "Cannot create the output folder:" & vbCrLf & outFolder, vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    m_logPath = ""
    If WRITE_LOG Then
        m_logPath = m_fso.BuildPath(outFolder, _
                    "PrintCorrectSize_" & SanitizeFileName(m_fso.GetBaseName(df.FullName)) & ".log")
        On Error Resume Next
        m_fso.CreateTextFile m_logPath, True
        On Error GoTo 0
        Err.Clear
    End If

    Set originalModel = ActiveModelReference
    BuildPrintList models, printModels, n

    If n = 0 Then
        MsgBox "No printable models found.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    ReDim infos(0 To n - 1)
    ReDim names(0 To n - 1)
    ReDim paths(0 To n - 1)

    ' Scan models. Activation is needed for the selection set and for a
    ' reliable per-model scan in V8i.
    For i = 1 To n
        Set m = printModels(i - 1)
        m.Activate
        names(i - 1) = m.Name
        GetPageInfo m, pi
        infos(i - 1) = pi
        paths(i - 1) = MakePdfPath(df, m, n, outFolder)
    Next i
    originalModel.Activate

    ' Build a short report and ask for confirmation.
    summary = "Active DGN: " & df.FullName & vbCrLf & _
              "Output: " & outFolder & vbCrLf & vbCrLf

    For i = 0 To n - 1
        If infos(i).Valid Then
            summary = summary & Format$(i + 1, "00") & ". " & names(i) & _
                      "  ->  " & Format$(infos(i).WidthMM, "0.0") & " x " & _
                      Format$(infos(i).HeightMM, "0.0") & " mm  [" & infos(i).Source & "]"
            If infos(i).Paper <> "CUSTOM" Then
                summary = summary & "  ->  " & infos(i).Paper & _
                          " (measured " & Format$(infos(i).MeasuredWidthMM, "0.0") & " x " & _
                          Format$(infos(i).MeasuredHeightMM, "0.0") & " mm)"
            End If
            summary = summary & vbCrLf
        Else
            summary = summary & Format$(i + 1, "00") & ". " & names(i) & _
                      "  ->  SKIPPED (no usable geometry)" & vbCrLf
        End If
    Next i

    summary = summary & vbCrLf & "Create these PDF files?"
    answer = MsgBox(summary, vbOKCancel + vbInformation, "PrintCorrectSize")
    If answer <> vbOK Then Exit Sub

    ' One-time print setup.
    CadInputQueue.SendKeyin "PRINT DRIVER " & PDF_DRIVER
    CadInputQueue.SendKeyin "PRINT DESTINATION PLOTFILE"
    CadInputQueue.SendKeyin "PRINT COLORMODE COLOR"
    CadInputQueue.SendKeyin "PRINT AUTOCENTER ON"
    ' Do not add MicroStation's own plot border or draw the fence boundary.
    CadInputQueue.SendKeyin "PRINT ATTRIBUTES PLOTBORDER OFF"
    CadInputQueue.SendKeyin "PRINT ATTRIBUTES FENCEBOUNDARY OFF"

    LogLine "==== " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===="
    LogLine "DGN: " & df.FullName
    LogLine "Driver: " & PDF_DRIVER

    For i = 0 To n - 1
        If infos(i).Valid Then
            Set m = printModels(i)
            m.Activate

            If PrintOneModel(m, infos(i), paths(i)) Then
                okCount = okCount + 1
                LogLine "OK   " & names(i) & " | " & _
                        Format$(infos(i).WidthMM, "0.0") & "x" & Format$(infos(i).HeightMM, "0.0") & _
                        " mm | " & infos(i).Paper & " | measured " & _
                        Format$(infos(i).MeasuredWidthMM, "0.0") & "x" & _
                        Format$(infos(i).MeasuredHeightMM, "0.0") & " mm | " & _
                        infos(i).Source & " | " & paths(i)
            Else
                failCount = failCount + 1
                LogLine "FAIL " & names(i) & " | " & paths(i)
            End If
        Else
            skipped = skipped + 1
            LogLine "SKIP " & names(i)
        End If
    Next i

    ' Leave the file as we found it.
    On Error Resume Next
    ActiveDesignFile.Fence.Undefine
    On Error GoTo 0
    originalModel.Activate

    result = "Done." & vbCrLf & vbCrLf & _
             "Created: " & okCount & vbCrLf & _
             "Failed:  " & failCount & vbCrLf & _
             "Skipped: " & skipped & vbCrLf & vbCrLf & _
             "Folder: " & outFolder

    If failCount > 0 Then
        result = result & vbCrLf & vbCrLf & _
                 "Check the PrintCorrectSize log and make sure the Bentley PDF driver " & _
                 "(" & PDF_DRIVER & ") is installed and allows paper size editing."
    End If

    MsgBox result, IIf(failCount > 0, vbExclamation, vbInformation), "PrintCorrectSize"
End Sub

' Shows the detected page sizes without printing.
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
        MsgBox "Open a DGN file first.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set df = ActiveDesignFile
    Set models = df.Models
    If models.Count = 0 Then
        MsgBox "The active DGN has no models.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    Set originalModel = ActiveModelReference
    BuildPrintList models, printModels, n

    If n = 0 Then
        MsgBox "No printable models found.", vbExclamation, "PrintCorrectSize"
        Exit Sub
    End If

    report = "Active DGN: " & df.FullName & vbCrLf & vbCrLf

    For i = 1 To n
        Set m = printModels(i - 1)
        m.Activate
        GetPageInfo m, pi

        report = report & Format$(i, "00") & ". " & m.Name & vbCrLf
        If pi.Valid Then
            report = report & "     " & Format$(pi.WidthMM, "0.0") & " x " & _
                     Format$(pi.HeightMM, "0.0") & " mm  [" & pi.Source & "]"
            If pi.Paper <> "CUSTOM" Then
                report = report & "  ->  " & pi.Paper & _
                         " (measured " & Format$(pi.MeasuredWidthMM, "0.0") & " x " & _
                         Format$(pi.MeasuredHeightMM, "0.0") & " mm)"
            End If
            report = report & vbCrLf
        Else
            report = report & "     SKIPPED - no usable geometry" & vbCrLf
        End If
    Next i
    originalModel.Activate

    MsgBox report, vbInformation, "PrintCorrectSize - preview"
End Sub

'==============================================================================
' Model selection
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
' Page-size detection
'==============================================================================

Private Sub GetPageInfo(ByVal m As ModelReference, ByRef pi As PageInfo)
    Dim sd As SheetDefinition
    Dim b As BBox

    pi.Valid = False
    pi.Source = ""
    pi.Paper = ""
    pi.IsSheet = (m.Type = msdModelTypeSheet)
    pi.Box.Valid = False

    ' 1. Sheet model: the sheet definition is the authoritative page size.
    If pi.IsSheet Then
        On Error Resume Next
        Err.Clear
        Set sd = m.GetSheetDefinition
        If Err.Number = 0 Then
            If Not sd Is Nothing Then
                If sd.Width > 0# And sd.Height > 0# Then
                    SetPageSize pi, UORsToMM(m, sd.Width), UORsToMM(m, sd.Height), "SHEET"
                End If
            End If
        End If
        On Error GoTo 0
        Err.Clear
        If pi.Valid Then Exit Sub

        ' A sheet model without a usable sheet definition must be printed
        ' from a fence, like a normal design model.
        pi.IsSheet = False
    End If

    ' 2. Elements on a frame/plot level. This is the normal path for design
    '    models with a frame on M_RAMKA or _RAMKA.
    If Len(Trim$(BORDER_LEVEL_HINTS)) > 0 Then
        On Error Resume Next
        Err.Clear
        b = BBoxFromBorderLevels(m)
        If Err.Number = 0 And b.Valid Then
            SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                        UORsToMM(m, b.MaxY - b.MinY), "BORDER LEVEL"
            If pi.Valid Then
                pi.Box = b
                Exit Sub
            End If
        End If
        On Error GoTo 0
        Err.Clear
    End If

    ' 3. Manual override: selected elements.
    If USE_SELECTION_AS_PAGE Then
        On Error Resume Next
        Err.Clear
        If m.AnyElementsSelected Then
            b = BBoxFromEnumerator(m.GetSelectedElements)
            If b.Valid Then
                SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                            UORsToMM(m, b.MaxY - b.MinY), "SELECTION"
                If pi.Valid Then
                    pi.Box = b
                    Exit Sub
                End If
            End If
        End If
        On Error GoTo 0
        Err.Clear
    End If

    ' 4. Largest closed element (usually the sheet border).
    On Error Resume Next
    Err.Clear
    b = BBoxFromLargestClosedElement(m)
    If Err.Number = 0 And b.Valid Then
        SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                    UORsToMM(m, b.MaxY - b.MinY), "CLOSED ELEMENT"
        If pi.Valid Then
            pi.Box = b
            Exit Sub
        End If
    End If
    On Error GoTo 0
    Err.Clear

    ' 5. Fallback: all graphical elements.
    On Error Resume Next
    Err.Clear
    b = BBoxFromAllElements(m)
    If Err.Number = 0 And b.Valid Then
        SetPageSize pi, UORsToMM(m, b.MaxX - b.MinX), _
                    UORsToMM(m, b.MaxY - b.MinY), "EXTENTS"
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

                ' Text and dimensions on the frame level must not enlarge the
                ' page. They are usually notes placed outside the frame.
                If Not el.IsTextElement And Not el.IsTextNodeElement And _
                   Not el.IsDimensionElement Then
                    BBoxAdd allOnLevels, r
                End If

                ' Prefer the largest closed element (the frame shape) over the
                ' union of everything on the level. This ignores notes or other
                ' elements that happen to share the frame level.
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

    ' Use whichever is larger: the frame shape or the union of non-text
    ' elements on the frame level. A closed frame shape normally wins; when the
    ' frame is drawn from lines, the union wins over a smaller closed title
    ' block.
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
    ' Ignore null/invalid ranges.
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
    pi.Paper = "CUSTOM"

    If SnapToStandard(wMM, hMM, stdW, stdH, paperName) Then
        pi.WidthMM = stdW
        pi.HeightMM = stdH
        pi.Paper = paperName
    End If

    pi.Valid = IsSanePage(pi.WidthMM, pi.HeightMM)
End Sub

' Returns True when the measured page is within the tolerance of one of the
' sizes listed in STANDARD_SIZES_TO_SNAP. The returned width/height keep the
' orientation of the measured page.
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
    paperName = "CUSTOM"

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

' ISO A and B sizes in millimetres, portrait orientation.
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
' Printing
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

    ' Select the print area first. For a sheet model this also makes
    ' MicroStation use the sheet boundary.
    If pi.IsSheet Then
        CadInputQueue.SendKeyin "PRINT BOUNDARY SHEET"
    Else
        pageBox = pi.Box
        If Not DefineFenceForPage(m, pageBox) Then
            LogLine "  cannot define fence for " & m.Name
            Exit Function
        End If
        CadInputQueue.SendKeyin "PRINT BOUNDARY FENCE"
    End If

    ' A Bentley driver uses a portrait form plus an orientation. Set the
    ' short side as X and the long side as Y, then rotate if needed.
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
    ' Set the form while it is in portrait orientation, then rotate.
    CadInputQueue.SendKeyin "PRINT ORIENTATION PORTRAIT"
    CadInputQueue.SendKeyin "PRINT XFORM " & KeyinNumber(shortMM) & " MM"
    CadInputQueue.SendKeyin "PRINT YFORM " & KeyinNumber(longMM) & " MM"
    CadInputQueue.SendKeyin "PRINT ORIENTATION " & orientation
    CadInputQueue.SendKeyin "PRINT AUTOCENTER ON"
    CadInputQueue.SendKeyin "PRINT MAXIMIZE"

    ' Remove a stale file so a timeout cannot be mistaken for success.
    On Error Resume Next
    If m_fso.FileExists(pdfPath) Then m_fso.DeleteFile pdfPath, True
    On Error GoTo errHandler
    Err.Clear
    If m_fso.FileExists(pdfPath) Then
        LogLine "  cannot delete stale PDF: " & pdfPath
        Exit Function
    End If

    ' Use a short folder path and a file name without spaces. This avoids
    ' ambiguity when the key-in parser sees a path with spaces.
    keyinPath = MakeKeyinPath(pdfPath)
    CadInputQueue.SendKeyin "PRINT EXECUTE " & keyinPath

    PrintOneModel = WaitForPdf(pdfPath, PDF_TIMEOUT_SEC)
    Exit Function

errHandler:
    LogLine "  error " & Err.Number & ": " & Err.Description
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

                ' File size unchanged for ~0.6 s -> assume the PDF is complete.
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
        ' Midnight rollover.
        ElapsedSeconds = (86400# - started) + nowValue
    End If
End Function

'==============================================================================
' Files, names, small helpers
'==============================================================================

Private Function PrintDriverExists(ByVal driverName As String) As Boolean
    Dim paths As String
    Dim parts As Variant
    Dim i As Long
    Dim folder As String

    PrintDriverExists = False
    If Len(Trim$(driverName)) = 0 Then Exit Function

    ' Allow PDF_DRIVER to be a full path as well as a file name.
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

' Returns a path suitable for the key-in parser. The short folder path removes
' spaces from directory names; the file name is already sanitized by
' SanitizeFileName. If a space somehow survives, the path is quoted.
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
    Dim ch As String
    Dim result As String

    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", " ", ";")
    result = s

    For i = LBound(badChars) To UBound(badChars)
        result = Replace(result, CStr(badChars(i)), "_")
    Next i

    result = Trim$(result)
    If Len(result) = 0 Then result = "model"
    If Len(result) > 120 Then result = Left$(result, 120)
    SanitizeFileName = result
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
    Set ts = m_fso.OpenTextFile(m_logPath, 8, True)
    If Not ts Is Nothing Then
        ts.WriteLine text
        ts.Close
    End If
    On Error GoTo 0
End Sub
