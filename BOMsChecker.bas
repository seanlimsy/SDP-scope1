Option Explicit
Dim wbCompiled As Workbook
Dim wbWetCompiled As Worksheet
Dim wbDryCompiled As Worksheet

Dim wetSchedule As Worksheet
Dim dryScheduleD1 As Worksheet
Dim dryScheduleD2 As Worksheet
Dim dryScheduleD3 As Worksheet
Dim dryScheduleD4 As Worksheet
Dim missingBOMs As Worksheet

Sub checkBOMsMain()
    Dim nextStart As Integer
    initializeWorksheets
    nextStart = checkWetSchedule(wetSchedule, 1)
    nextStart = checkDrySchedule(dryScheduleD1, nextStart, "D1")
    nextStart = checkDrySchedule(dryScheduleD2, nextStart, "D2")
    nextStart = checkDrySchedule(dryScheduleD3, nextStart, "D3")
    nextStart = checkDrySchedule(dryScheduleD4, nextStart, "D4")
End Sub

Sub initializeWorksheets()
    Set wbCompiled = ThisWorkbook
    setWorksheet wbWetCompiled, "WetBOMs"
    setWorksheet wbDryCompiled, "DryBOMs"
    setWorksheet wetSchedule, "Baseline"
    setWorksheet dryScheduleD1, "Blender 1 Schedule"
    setWorksheet dryScheduleD2, "Blender 2 Schedule"
    setWorksheet dryScheduleD3, "Blender 3 Schedule"
    setWorksheet dryScheduleD4, "Blender 4 Schedule"

    wbCompiled.Sheets.Add.Name = "MissingBOMs"
    Set missingBOMs = wbCompiled.Sheets("MissingBOMs")
    setHeaderRows missingBOMs
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wbCompiled.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Sub setHeaderRows(ws)
    ws.Range("A1") = "Process Type"
    ws.Range("B1") = "Powder Code"
    ws.Range("C1") = "BOM Available?"
End Sub

Function checkWetSchedule(ws, startRow)
    Dim allBPListCodes As Range
    Dim uniqueBPListCodes As Variant
    Dim BPToCheck As String
    Dim i As Integer

    Dim wetBOMRange As Range
    Set allBPListCodes = ws.Range("B2:B" & ws.Range("B2").End(xlDown).Row)
    uniqueBPListCodes = WorksheetFunction.Unique(allBPListCodes)
    For i = LBound(uniqueBPListCodes) To UBound(uniqueBPListCodes)
        BPToCheck = uniqueBPListCodes(i, 1)
        missingBOMs.Range("A" & (i + startRow)) = "WP"
        missingBOMs.Range("B" & (i + startRow)) = BPToCheck
        missingBOMs.Range("C" & (i + startRow)) = isBPinBOMs(BPToCheck, wbWetCompiled)
    Next i
    
    Dim nextStartRow As Integer
    nextStartRow = missingBOMs.Range("A1").End(xlDown).Row
    checkWetSchedule = nextStartRow
End Function

Function checkDrySchedule(ws, startRow, Di)
    Dim allFPListCodes As Range
    Dim uniqueFPListCodes As Variant
    Dim FPToCheck As String
    Dim i As Integer

    Dim dryBOMRange As Range
    Set allFPListCodes = ws.Range("B2:B" & ws.Range("B2").End(xlDown).Row)
    uniqueFPListCodes = WorksheetFunction.Unique(allFPListCodes)
    For i = LBound(uniqueFPListCodes) To UBound(uniqueFPListCodes)
        FPToCheck = uniqueFPListCodes(i, 1)
        missingBOMs.Range("A" & (i + startRow)) = "DB-" & Di
        missingBOMs.Range("B" & (i + startRow)) = FPToCheck
        missingBOMs.Range("C" & (i + startRow)) = isBPinBOMs(FPToCheck, wbDryCompiled)
    Next i
    Dim nextStartRow As Integer
    nextStartRow = missingBOMs.Range("A1").End(xlDown).Row
    checkDrySchedule = nextStartRow
End Function

Function isBPinBOMs(powerCode, BOMws) As Boolean
    Dim firstRow As Integer, lastRow As Integer
    firstRow = 2
    lastRow = BOMws.Range("B2").End(xlDown).Row

    Dim i As Integer
    For i = firstRow To lastRow
        If BOMws.Range("B" & i) = powerCode Then
            isBPinBOMs = True
            Exit Function
        End If
    Next i
    isBPinBOMs = False
End Function

