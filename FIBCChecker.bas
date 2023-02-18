Option Explicit
Dim wbCompiled As Workbook
Dim wbWetCompiled As Worksheet
Dim wbDryCompiled As Worksheet
Dim FIBCMaterialsWS As Worksheet
Dim missingFIBCs As Worksheet

Sub checkFIBCsMain()
    Dim nextStart As Integer
    initializeWorksheets
    checkBOMs wbWetCompiled, "Wet"
    checkBOMs wbDryCompiled, "Dry"

End Sub

Sub initializeWorksheets()
    Set wbCompiled = ThisWorkbook
    setWorksheet wbWetCompiled, "WetBOMs"
    setWorksheet wbDryCompiled, "DryBOMs"
    setWorksheet FIBCMaterialsWS, "FIBC Materials"

    setHeaderRows FIBCMaterialsWS
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
    ws.Range("H1") = "Handling Type = Metered in Wet?"
    ws.Range("I1") = "Handling Type = Metered in Dry?"
End Sub

Sub checkBOMs(ws, process)
    Dim firstFIBCRow As Integer, lastFIBCRow As Integer    
    firstFIBCRow = 2
    lastFIBCRow = FIBCMaterialsWS.Range("A1").End(xlDown).Row

    Dim wsLastRow As Integer
    wsLastRow = ws.Range("C1").End(xlDown).Row

    Dim colLetter As String
    If process = "Wet" Then 
        colLetter = "H"
    ElseIf process = "Dry" Then 
        colLetter = "I"
    End If

    Dim i as Integer
    Dim FIBCCode As String
    Dim NFIBCCodes As Integer, NFIBCCodesAndMetered As Integer
    For i = firstFIBCRow To lastFIBCRow
        FIBCCode = FIBCMaterialsWS.Range("A" & i).Value
        NFIBCCodes = WorksheetFunction.CountIf(ws.Range("C2:C" & wsLastRow), "=" & FIBCCode)
        NFIBCCodesAndMetered = WorksheetFunction.CountIfs(ws.Range("C2:C" & wsLastRow), "=" & FIBCCode, ws.Range("H2:H" & wsLastRow), "=Metered")

        If NFIBCCodes = 0 Then 
            FIBCMaterialsWS.Range(colLetter & i) = "Not in Processes"
        ElseIf NFIBCCodes = NFIBCCodesAndMetered Then
            FIBCMaterialsWS.Range(colLetter & i) = "True"
        Else
            FIBCMaterialsWS.Range(colLetter & i) = "False"
        End If
    Next i

End Sub