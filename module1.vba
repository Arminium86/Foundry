Option Explicit

'================================================================================
' MAIN
'================================================================================
Public Sub ExportCombinedCSV_WithOPFCode_Scenario_DynamicMapping()

    Dim wsList As Worksheet, ws As Worksheet
    Dim rng As Range, c As Range
    Dim shName As String

    Set wsList = ActiveSheet
    Set rng = wsList.Range("A2:A11")   ' sheet names (controller)

    ' ---- Build OPF -> OPF Code mapping from C3:D20 ----
    Dim dictMap As Object
    Set dictMap = CreateObject("Scripting.Dictionary")
    dictMap.CompareMode = vbTextCompare

    Dim mapCell As Range
    For Each mapCell In wsList.Range("C3:C20").Cells
        If Trim$(CStr(mapCell.Value2)) <> "" And Trim$(CStr(mapCell.Offset(0, 1).Value2)) <> "" Then
            dictMap(Trim$(CStr(mapCell.Value2))) = Trim$(CStr(mapCell.Offset(0, 1).Value2))
        End If
    Next mapCell

    ' Track scenarios actually combined
    Dim scenUsed As Object
    Set scenUsed = CreateObject("Scripting.Dictionary")
    scenUsed.CompareMode = vbTextCompare

    ' Speed
    Dim calcMode As XlCalculation, evtMode As Boolean
    Application.ScreenUpdating = False
    evtMode = Application.EnableEvents: Application.EnableEvents = False
    calcMode = Application.Calculation: Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    ' Temp workbook (export only)
    Dim wbTmp As Workbook, wsTmp As Worksheet
    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    Set wsTmp = wbTmp.Worksheets(1)
    wsTmp.Cells.Clear

    Dim outRow As Long: outRow = 1
    Dim headerWritten As Boolean
    Dim headerCols As Long

    ' ---- 1) COMBINE DATA + SCENARIO (STRICT header width based on VALUES ONLY) ----
    For Each c In rng.Cells
        shName = Trim$(CStr(c.Value2))
        If shName = "" Then GoTo NextName

        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(shName)
        On Error GoTo CleanUp
        If ws Is Nothing Then GoTo NextName

        ' Strict: last header col with a VALUE in row 1 (ignores formatting/phantom used range)
        Dim thisHeaderCols As Long
        thisHeaderCols = LastHeaderColByValues(ws, 1)
        If thisHeaderCols = 0 Then GoTo NextName

        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow < 2 Then GoTo NextName

        If Not headerWritten Then
            headerCols = thisHeaderCols

            ' headers
            wsTmp.Range(wsTmp.Cells(outRow, 1), wsTmp.Cells(outRow, headerCols)).Value2 = _
                ws.Range(ws.Cells(1, 1), ws.Cells(1, headerCols)).Value2

            wsTmp.Cells(outRow, headerCols + 1).Value2 = "Scenario"
            outRow = outRow + 1
            headerWritten = True
        End If

        ' If other sheets have different widths, we still only take the first-sheet width
        Dim takeCols As Long
        takeCols = headerCols
        If thisHeaderCols < takeCols Then takeCols = thisHeaderCols

        Dim dataArr As Variant, scenArr() As Variant, i As Long
        dataArr = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, takeCols)).Value2
        ReDim scenArr(1 To UBound(dataArr, 1), 1 To 1)

        For i = 1 To UBound(scenArr, 1)
            scenArr(i, 1) = shName
        Next i

        wsTmp.Range(wsTmp.Cells(outRow, 1), wsTmp.Cells(outRow + UBound(dataArr, 1) - 1, takeCols)).Value2 = dataArr
        wsTmp.Range(wsTmp.Cells(outRow, headerCols + 1), wsTmp.Cells(outRow + UBound(scenArr, 1) - 1, headerCols + 1)).Value2 = scenArr

        outRow = outRow + UBound(dataArr, 1)

        If Not scenUsed.Exists(shName) Then scenUsed(shName) = True

NextName:
        Set ws = Nothing
    Next c

    If Not headerWritten Then Err.Raise vbObjectError + 1, , "No valid sheets found."

    Dim finalLastRow As Long
    finalLastRow = wsTmp.Cells(wsTmp.Rows.Count, 1).End(xlUp).Row

    ' ---- 2) ADD OPF CODE ----
    Dim opfCol As Long, codeCol As Long
    opfCol = FindHeaderColumn(wsTmp, "OPF")
    If opfCol = 0 Then Err.Raise vbObjectError + 2, , "'OPF' column not found."

    codeCol = AddOrGetHeaderColumn(wsTmp, "OPF Code")

    Dim inArr As Variant, outArr() As Variant, r As Long
    inArr = wsTmp.Range(wsTmp.Cells(2, opfCol), wsTmp.Cells(finalLastRow, opfCol)).Value2
    ReDim outArr(1 To UBound(inArr, 1), 1 To 1)

    For r = 1 To UBound(inArr, 1)
        Dim opfVal As String
        opfVal = Trim$(CStr(inArr(r, 1)))
        If opfVal = "" Then
            outArr(r, 1) = "blank"
        ElseIf dictMap.Exists(opfVal) Then
            outArr(r, 1) = dictMap(opfVal)
        Else
            outArr(r, 1) = vbNullString
        End If
    Next r

    wsTmp.Range(wsTmp.Cells(2, codeCol), wsTmp.Cells(finalLastRow, codeCol)).Value2 = outArr

    ' ---- 3) ADD brand2 and port_tw@f2 (duplicates) ----
    Dim colBrand As Long, colBrand2 As Long
    Dim colTwF As Long, colTwF2 As Long
    Dim colTwL As Long, colTwL2 As Long
    Dim colYear As Long
    
    colBrand = FindHeaderColumn(wsTmp, "brand")
    colTwF = FindHeaderColumn(wsTmp, "port_tw@f")
    colTwL = FindHeaderColumn(wsTmp, "port_tw@l")
    colYear = FindHeaderColumn(wsTmp, "year")
    
    If colBrand = 0 Then Err.Raise vbObjectError + 10, , "'brand' column not found."
    If colTwF = 0 Then Err.Raise vbObjectError + 11, , "'port_tw@f' column not found."
    If colTwL = 0 Then Err.Raise vbObjectError + 13, , "'port_tw@l' column not found."
    If colYear = 0 Then Err.Raise vbObjectError + 12, , "'year' column not found."
    
    colBrand2 = AddOrGetHeaderColumn(wsTmp, "brand2")
    colTwF2 = AddOrGetHeaderColumn(wsTmp, "port_tw@f2")
    colTwL2 = AddOrGetHeaderColumn(wsTmp, "port_tw@l2")
    
    For r = 2 To finalLastRow
        wsTmp.Cells(r, colBrand2).Value2 = wsTmp.Cells(r, colBrand).Value2
        wsTmp.Cells(r, colTwF2).Value2 = wsTmp.Cells(r, colTwF).Value2
        wsTmp.Cells(r, colTwL2).Value2 = wsTmp.Cells(r, colTwL).Value2
    Next r


    ' ---- 4) Apply FL fines allocation to brand2 ----
    ApplyFLFinesAllocationToBrand2 ThisWorkbook, wsTmp, colBrand2, colYear, 2, finalLastRow

    ' ---- 5) Populate Buffer_Abs_Export from results ----
    PopulateBufferAbsExportFromResults ThisWorkbook, wsTmp

    ' ---- 6) Append Buffer_Abs_Export (scenario-specific) ----
    AppendBufferAbsExportPerScenario ThisWorkbook, wsTmp, scenUsed, dictMap

    ' Recalculate last row and re-apply FL mapping (covers appended rows too)
    Dim lastRowAfterAppend As Long
    lastRowAfterAppend = wsTmp.Cells(wsTmp.Rows.Count, 1).End(xlUp).Row
    If lastRowAfterAppend > finalLastRow Then
        ApplyFLFinesAllocationToBrand2 ThisWorkbook, wsTmp, colBrand2, colYear, finalLastRow + 1, lastRowAfterAppend
    End If

    ' ---- 7) Export CSV ----
    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=GetDesktopPath() & "\Combined_" & Format(Now, "yyyymmdd_hhnnss") & ".csv", _
        FileFilter:="CSV UTF-8 (*.csv), *.csv", _
        Title:="Save combined CSV to Desktop" _
    )

    If savePath <> False Then
        Application.DisplayAlerts = False
        wbTmp.SaveAs Filename:=CStr(savePath), FileFormat:=xlCSVUTF8
        Application.DisplayAlerts = True
    End If

    wbTmp.Close SaveChanges:=False

CleanUp:
    Application.Calculation = calcMode
    Application.EnableEvents = evtMode
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "Export failed: " & Err.Description, vbExclamation
    End If
End Sub

'================================================================================
' Buffer build: calculate absolute buffer rows from results by scenario
'================================================================================
Private Sub PopulateBufferAbsExportFromResults(ByVal wb As Workbook, ByVal wsCombined As Worksheet)
    Const SH_BUF As String = "Buffers & OPF Physical Caps"
    Const TBL_BUF As String = "Buffer_Abs_Export"

    Dim colScenario As Long, colOpfCode As Long, colBrand As Long, colBrand2 As Long, colYear As Long
    Dim colTwF As Long, colTwL As Long

    colScenario = FindHeaderColumn(wsCombined, "Scenario")
    colOpfCode = FindHeaderColumn(wsCombined, "OPF Code")
    colBrand = FindHeaderColumn(wsCombined, "brand")
    colBrand2 = FindHeaderColumn(wsCombined, "brand2")
    colYear = FindHeaderColumn(wsCombined, "year")
    colTwF = FindHeaderColumn(wsCombined, "port_tw@f")
    colTwL = FindHeaderColumn(wsCombined, "port_tw@l")

    If colScenario = 0 Or colOpfCode = 0 Or colBrand = 0 Or colBrand2 = 0 Or colYear = 0 Or colTwF = 0 Or colTwL = 0 Then
        Err.Raise vbObjectError + 150, , "Combined sheet missing required columns for abs buffer calc."
    End If

    Dim lastRow As Long
    lastRow = wsCombined.Cells(wsCombined.Rows.Count, colScenario).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim scenArr As Variant, opfArr As Variant, brandArr As Variant, brand2Arr As Variant
    Dim yearArr As Variant, twFArr As Variant, twLArr As Variant
    scenArr = wsCombined.Range(wsCombined.Cells(2, colScenario), wsCombined.Cells(lastRow, colScenario)).Value2
    opfArr = wsCombined.Range(wsCombined.Cells(2, colOpfCode), wsCombined.Cells(lastRow, colOpfCode)).Value2
    brandArr = wsCombined.Range(wsCombined.Cells(2, colBrand), wsCombined.Cells(lastRow, colBrand)).Value2
    brand2Arr = wsCombined.Range(wsCombined.Cells(2, colBrand2), wsCombined.Cells(lastRow, colBrand2)).Value2
    yearArr = wsCombined.Range(wsCombined.Cells(2, colYear), wsCombined.Cells(lastRow, colYear)).Value2
    twFArr = wsCombined.Range(wsCombined.Cells(2, colTwF), wsCombined.Cells(lastRow, colTwF)).Value2
    twLArr = wsCombined.Range(wsCombined.Cells(2, colTwL), wsCombined.Cells(lastRow, colTwL)).Value2

    Dim sums As Object
    Set sums = CreateObject("Scripting.Dictionary")
    sums.CompareMode = vbTextCompare

    Dim r As Long
    For r = 1 To UBound(scenArr, 1)
        Dim scen As String
        scen = Trim$(CStr(scenArr(r, 1)))
        If scen = "" Then GoTo NextRow

        Dim opf As String
        opf = Trim$(CStr(opfArr(r, 1)))
        If opf = "" Or UCase$(opf) = "BLANK" Then GoTo NextRow

        Dim br As String
        br = UCase$(Trim$(CStr(brandArr(r, 1))))
        If br = "" Then GoTo NextRow

        Dim br2 As String
        br2 = UCase$(Trim$(CStr(brand2Arr(r, 1))))
        If br2 = "" Then GoTo NextRow

        Dim per As Long
        per = CLng(Val(yearArr(r, 1)))
        If per <= 0 Then GoTo NextRow

        If br = "WASTE" Or br2 = "WASTE" Then GoTo NextRow

        If br <> "FL" And br2 <> br Then GoTo NextRow

        Dim k As String
        Dim baseVal As Double

        If br = "FL" Then
            baseVal = NzNum(twLArr(r, 1))
            k = scen & "|" & opf & "|" & br & "|" & CStr(per)
            If sums.Exists(k) Then
                sums(k) = CDbl(sums(k)) + baseVal
            Else
                sums(k) = baseVal
            End If

            If br2 <> "FL" Then
                baseVal = NzNum(twFArr(r, 1))
                k = scen & "|" & opf & "|" & br2 & "|" & CStr(per)
                If sums.Exists(k) Then
                    sums(k) = CDbl(sums(k)) + baseVal
                Else
                    sums(k) = baseVal
                End If
            End If
        Else
            baseVal = NzNum(twFArr(r, 1))
            k = scen & "|" & opf & "|" & br & "|" & CStr(per)
            If sums.Exists(k) Then
                sums(k) = CDbl(sums(k)) + baseVal
            Else
                sums(k) = baseVal
            End If
        End If
NextRow:
    Next r

    If sums.Count = 0 Then Exit Sub

    Dim closingBuf As Object
    Set closingBuf = LoadClosingBuffers(wb)

    Dim addBuf As Object
    Set addBuf = LoadAdditionalOpfBuffer(wb)

    Dim perfPct As Double
    perfPct = NormalizePct(LoadPerfectSolveBuffer(wb))

    Dim wsBuf As Worksheet
    On Error Resume Next
    Set wsBuf = wb.Worksheets(SH_BUF)
    On Error GoTo 0
    If wsBuf Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsBuf.ListObjects(TBL_BUF)
    On Error GoTo 0

    If lo Is Nothing Then
        Dim startCell As Range
        Set startCell = wsBuf.Range("A1")
        startCell.Resize(1, 5).Value = Array("OPF_Bucket", "Brand", "Foundry_Period", "AbsBufferTonnes", "Scenario")
        Set lo = wsBuf.ListObjects.Add(xlSrcRange, startCell.Resize(2, 5), , xlYes)
        lo.Name = TBL_BUF
    Else
        EnsureTableColumn lo, "Scenario"
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    End If

    Dim cOPF As Long, cBrand As Long, cPer As Long, cAbs As Long, cScenario As Long
    cOPF = GetLoColIndexAny(lo, Array("OPF", "OPF_Bucket"))
    cBrand = GetLoColIndexAny(lo, Array("Brand"))
    cPer = GetLoColIndexAny(lo, Array("Period", "Foundry_Period", "FOUNDRY_PERIOD"))
    cAbs = GetLoColIndexAny(lo, Array("AbsBufferTonnes"))
    cScenario = GetLoColIndexAny(lo, Array("Scenario"))

    Dim outArr() As Variant
    ReDim outArr(1 To sums.Count, 1 To lo.ListColumns.Count)

    Dim i As Long
    i = 0
    Dim key As Variant
    For Each key In sums.Keys
        Dim parts() As String
        parts = Split(CStr(key), "|")

        Dim opfKey As String
        opfKey = UCase$(parts(1))
        If opfKey = "" Then GoTo NextKey

        i = i + 1
        Dim brandKey As String
        brandKey = UCase$(parts(2))
        Dim periodKey As Long
        periodKey = CLng(parts(3))

        Dim addPct As Double
        addPct = 0#
        If addBuf.Exists(opfKey) Then addPct = NormalizePct(addBuf(opfKey))

        Dim baseVal2 As Double
        baseVal2 = CDbl(sums(key))

        Dim absTotal As Double
        If periodKey = 1 Then
            Dim closeAbs As Double
            closeAbs = 0#
            Dim closeKey As String
            closeKey = MapClosingBufferOpf(opfKey) & brandKey
            If closingBuf.Exists(closeKey) Then closeAbs = CDbl(closingBuf(closeKey))

            absTotal = closeAbs + (baseVal2 * addPct)
        Else
            absTotal = (baseVal2 * perfPct) + (baseVal2 * addPct)
        End If
        absTotal = absTotal * -1#

        outArr(i, cScenario) = parts(0)
        outArr(i, cOPF) = opfKey
        outArr(i, cBrand) = brandKey
        outArr(i, cPer) = periodKey
        outArr(i, cAbs) = absTotal
NextKey:
    Next key

    If i = 0 Then Exit Sub

    lo.HeaderRowRange.Offset(1, 0).Resize(i, lo.ListColumns.Count).Value = outArr
    lo.ListColumns(cAbs).DataBodyRange.NumberFormat = "#,##0"
End Sub

'================================================================================
' Buffer append: Buffer_Abs_Export appended per scenario actually combined
' Mapping:
'   OPF/OPF_Bucket -> OPF Code (via dictMap)
'   Brand          -> brand and brand2
'   Period         -> year
'   AbsBufferTonnes-> port_tw@f2
'   Scenario       -> Scenario
'================================================================================
Private Sub AppendBufferAbsExportPerScenario(ByVal wb As Workbook, ByVal wsCombined As Worksheet, _
                                            ByVal scenUsed As Object, ByVal dictMap As Object)

    Const SH_BUF As String = "Buffers & OPF Physical Caps"
    Const TBL_BUF As String = "Buffer_Abs_Export"

    Dim wsBuf As Worksheet
    On Error Resume Next
    Set wsBuf = wb.Worksheets(SH_BUF)
    On Error GoTo 0
    If wsBuf Is Nothing Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsBuf.ListObjects(TBL_BUF)
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub

    Dim colScenario As Long, colOpfCode As Long, colBrand As Long, colBrand2 As Long, colYear As Long
    Dim colTwF2 As Long, colTwL2 As Long
    
    colScenario = FindHeaderColumn(wsCombined, "Scenario")
    colOpfCode = FindHeaderColumn(wsCombined, "OPF Code")
    colBrand = FindHeaderColumn(wsCombined, "brand")
    colBrand2 = FindHeaderColumn(wsCombined, "brand2")
    colYear = FindHeaderColumn(wsCombined, "year")
    colTwF2 = FindHeaderColumn(wsCombined, "port_tw@f2")
    colTwL2 = FindHeaderColumn(wsCombined, "port_tw@l2")
    
    If colScenario = 0 Or colOpfCode = 0 Or colBrand = 0 Or colBrand2 = 0 Or colYear = 0 Or colTwF2 = 0 Or colTwL2 = 0 Then
        Err.Raise vbObjectError + 200, , "Combined sheet missing required columns for buffer append (need port_tw@f2 and port_tw@l2)."
    End If


    Dim arr As Variant
    arr = lo.DataBodyRange.Value2

    Dim cScenario As Long, cOPF As Long, cBrand As Long, cPer As Long, cAbs As Long
    cScenario = GetLoColIndexAny(lo, Array("Scenario"))
    cOPF = GetLoColIndexAny(lo, Array("OPF", "OPF_Bucket"))
    cBrand = GetLoColIndexAny(lo, Array("Brand"))
    cPer = GetLoColIndexAny(lo, Array("Period", "Foundry_Period", "FOUNDRY_PERIOD"))
    cAbs = GetLoColIndexAny(lo, Array("AbsBufferTonnes"))

    Dim bufRows As Long
    bufRows = UBound(arr, 1)
    If bufRows <= 0 Then Exit Sub

    Dim lastCol As Long
    lastCol = wsCombined.Cells(1, wsCombined.Columns.Count).End(xlToLeft).Column

    Dim scen As Variant
    For Each scen In scenUsed.Keys

        Dim startRow As Long
        startRow = wsCombined.Cells(wsCombined.Rows.Count, colScenario).End(xlUp).Row + 1

        Dim matched As Long
        matched = 0
        Dim r As Long
        For r = 1 To bufRows
            If StrComp(CStr(arr(r, cScenario)), CStr(scen), vbTextCompare) = 0 Then
                If Trim$(CStr(arr(r, cOPF))) <> "" Then
                    matched = matched + 1
                End If
            End If
        Next r
        If matched = 0 Then GoTo NextScenario

        ' Build a block (matched x lastCol) so we write once
        Dim outBlock() As Variant
        ReDim outBlock(1 To matched, 1 To lastCol)

        Dim outRow As Long
        outRow = 0

        For r = 1 To bufRows
            If StrComp(CStr(arr(r, cScenario)), CStr(scen), vbTextCompare) <> 0 Then GoTo NextBufRow

            Dim opfRaw As String
            opfRaw = Trim$(CStr(arr(r, cOPF)))
            If opfRaw = "" Then GoTo NextBufRow

            Dim opfCodeVal As String
            If dictMap.Exists(opfRaw) Then
                opfCodeVal = CStr(dictMap(opfRaw))
            Else
                ' If it isn't in mapping, keep the raw bucket (better than blank for diagnostics)
                opfCodeVal = opfRaw
            End If

            Dim br As String
            br = UCase$(Trim$(CStr(arr(r, cBrand))))

            Dim per As Long
            per = CLng(Val(arr(r, cPer)))

            Dim absT As Double
            absT = CDbl(NzNum(arr(r, cAbs)))

            outRow = outRow + 1
            outBlock(outRow, colScenario) = CStr(scen)
            outBlock(outRow, colOpfCode) = opfCodeVal
            outBlock(outRow, colBrand) = br
            outBlock(outRow, colBrand2) = br
            outBlock(outRow, colYear) = per
            If br = "FL" Then
                outBlock(outRow, colTwL2) = absT   ' exception: FL buffer rows go to port_tw@l2
            Else
                outBlock(outRow, colTwF2) = absT   ' default: everyone else goes to port_tw@f2
            End If
NextBufRow:
        Next r

        wsCombined.Range(wsCombined.Cells(startRow, 1), wsCombined.Cells(startRow + matched - 1, lastCol)).Value2 = outBlock
NextScenario:
    Next scen
End Sub

'================================================================================
' FL fines allocation: rename brand2="FL" per period(year) using FL_Fines_Allocation
'================================================================================
Private Sub ApplyFLFinesAllocationToBrand2(ByVal wb As Workbook, ByVal ws As Worksheet, _
                                          ByVal colBrand2 As Long, ByVal colYear As Long, _
                                          ByVal startRow As Long, ByVal endRow As Long)

    Dim alloc As Object
    Set alloc = LoadFLFinesAllocation(wb)

    If alloc Is Nothing Or alloc.Count = 0 Then Exit Sub
    If endRow < startRow Then Exit Sub

    Dim r As Long
    For r = startRow To endRow
        If UCase$(Trim$(CStr(ws.Cells(r, colBrand2).Value2))) = "FL" Then
            Dim perKey As String
            perKey = CStr(CLng(Val(ws.Cells(r, colYear).Value2)))
            If alloc.Exists(perKey) Then
                ws.Cells(r, colBrand2).Value2 = alloc(perKey)
            End If
        End If
    Next r
End Sub

Private Function LoadFLFinesAllocation(ByVal wb As Workbook) As Object
    Const SH As String = "Buffers & OPF Physical Caps"
    Const TBL As String = "FL_Fines_Allocation"

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = wb.Worksheets(SH)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(TBL)
    On Error GoTo 0

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        Set LoadFLFinesAllocation = dict
        Exit Function
    End If

    Dim arr As Variant
    arr = lo.DataBodyRange.Value2

    Dim cPer As Long, cAlloc As Long
    cPer = GetLoColIndexAny(lo, Array("Period"))
    cAlloc = GetLoColIndexAny(lo, Array("FL Fines Allocation", "FL_Fines_Allocation"))

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim per As Long: per = CLng(Val(arr(r, cPer)))
        Dim b As String: b = UCase$(Trim$(CStr(arr(r, cAlloc))))
        If per >= 1 And b <> "" Then dict(CStr(per)) = b
    Next r

    Set LoadFLFinesAllocation = dict
End Function

'================================================================================
' Buffers lookups
'================================================================================
Private Function LoadClosingBuffers(ByVal wb As Workbook) As Object
    Const SH As String = "Buffers & OPF Physical Caps"
    Const TBL As String = "Closing_Feedable_Stocks_Buffer"

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = wb.Worksheets(SH)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(TBL)
    On Error GoTo 0

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        Set LoadClosingBuffers = dict
        Exit Function
    End If

    Dim arr As Variant
    arr = lo.DataBodyRange.Value2

    Dim cProd As Long, cBuf As Long
    cProd = GetLoColIndexAny(lo, Array("Product"))
    cBuf = GetLoColIndexAny(lo, Array("Period 1 Closing Feedable Stocks Buffer"))

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim prod As String
        prod = UCase$(Trim$(CStr(arr(r, cProd))))
        If prod <> "" Then dict(prod) = NzNum(arr(r, cBuf))
    Next r

    Set LoadClosingBuffers = dict
End Function

Private Function LoadAdditionalOpfBuffer(ByVal wb As Workbook) As Object
    Const SH As String = "Buffers & OPF Physical Caps"
    Const TBL As String = "Additional_OPF_Buffer"

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = wb.Worksheets(SH)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(TBL)
    On Error GoTo 0

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        Set LoadAdditionalOpfBuffer = dict
        Exit Function
    End If

    Dim arr As Variant
    arr = lo.DataBodyRange.Value2

    Dim cOpf As Long, cBuf As Long
    cOpf = GetLoColIndexAny(lo, Array("OPF"))
    cBuf = GetLoColIndexAny(lo, Array("Additional OPF Buffer"))

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim opf As String
        opf = UCase$(Trim$(CStr(arr(r, cOpf))))
        If opf <> "" Then dict(opf) = NzNum(arr(r, cBuf))
    Next r

    Set LoadAdditionalOpfBuffer = dict
End Function

Private Function LoadPerfectSolveBuffer(ByVal wb As Workbook) As Double
    Const SH As String = "Buffers & OPF Physical Caps"
    Const TBL As String = "Perfect_Solve_Buffer"

    Dim ws As Worksheet, lo As ListObject
    On Error Resume Next
    Set ws = wb.Worksheets(SH)
    If Not ws Is Nothing Then Set lo = ws.ListObjects(TBL)
    On Error GoTo 0

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    LoadPerfectSolveBuffer = NzNum(lo.DataBodyRange.Cells(1, 1).Value2)
End Function

'================================================================================
' Utilities
'================================================================================
Private Sub EnsureTableColumn(ByVal lo As ListObject, ByVal headerText As String)
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).Name), headerText, vbTextCompare) = 0 Then Exit Sub
    Next i

    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.Count).Name = headerText
End Sub

Private Function NormalizePct(ByVal v As Variant) As Double
    Dim pct As Double
    pct = NzNum(v)
    If pct > 1# Then pct = pct / 100#
    NormalizePct = pct
End Function

Private Function MapClosingBufferOpf(ByVal opfKey As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(opfKey))

    Select Case normalized
        Case "VK"
            MapClosingBufferOpf = "KG"
        Case "EW"
            MapClosingBufferOpf = "EL"
        Case Else
            MapClosingBufferOpf = normalized
    End Select
End Function

' Values-only "last header col" (fixes phantom / formatting / table expansion issues)
Private Function LastHeaderColByValues(ByVal ws As Worksheet, ByVal headerRow As Long) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Rows(headerRow).Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If f Is Nothing Then
        LastHeaderColByValues = 0
    Else
        LastHeaderColByValues = f.Column
    End If
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, col As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For col = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, col).Value2)), headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = col
            Exit Function
        End If
    Next col
End Function

Private Function AddOrGetHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim col As Long
    col = FindHeaderColumn(ws, headerText)
    If col <> 0 Then
        AddOrGetHeaderColumn = col
        Exit Function
    End If

    col = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, col).Value2 = headerText
    AddOrGetHeaderColumn = col
End Function

Private Function GetDesktopPath() As String
    GetDesktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function

Private Function GetLoColIndexAny(ByVal lo As ListObject, ByVal headers As Variant) As Long
    Dim i As Long, h As Variant
    For Each h In headers
        For i = 1 To lo.ListColumns.Count
            If StrComp(Trim$(lo.ListColumns(i).Name), CStr(h), vbTextCompare) = 0 Then
                GetLoColIndexAny = i
                Exit Function
            End If
        Next i
    Next h
    Err.Raise vbObjectError + 555, , "None of the columns found in table '" & lo.Name & "': " & Join(headers, ", ")
End Function

Private Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Then
        NzNum = 0#
    ElseIf IsEmpty(v) Or v = vbNullString Then
        NzNum = 0#
    ElseIf IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = 0#
    End If
End Function
