Option Explicit
Dim wbCompiled As Workbook
Dim wbWetCompiled As Worksheet
Dim wbDryCompiled As Worksheet
Dim wbDryWetCompiled As Worksheet

Sub main()
    Dim wbCompiledPath As String
    wbCompiledPath = defineFiles()
    getWetBOMs wbCompiledPath
    getDryBOMs wbCompiledPath
    matchDryOnWet
    joinWetOnDry
    getFormattedMaterialCode wbWetCompiled
    getFormattedMaterialCode wbDryCompiled
    getFormattedMaterialCode wbDryWetCompiled
End Sub

Function defineFiles()
    Set wbCompiled = ThisWorkbook
    defineFiles = wbCompiled.Path
    wbCompiled.Sheets.Add.Name = "WetBOMs"
    Set wbWetCompiled = wbCompiled.Worksheets("WetBOMs")
    wbWetCompiled.Tab.Color = RGB(255, 192, 0)
    setHeaderRows wbWetCompiled

    wbCompiled.Sheets.Add.Name = "DryBOMs"
    Set wbDryCompiled = wbCompiled.Worksheets("DryBOMs")
    wbDryCompiled.Tab.Color = RGB(255, 192, 0)
    setHeaderRows wbDryCompiled

    wbCompiled.Sheets.Add.Name = "DryMerged"
    Set wbDryWetCompiled = wbCompiled.Worksheets("DryMerged")
    wbDryWetCompiled.Tab.Color = RGB(255, 192, 0)
    setHeaderRows wbDryWetCompiled
End Function

Sub setHeaderRows(ws)
    ws.Range("A1") = "Process"
    ws.Range("B1") = "Powder Code"
    ws.Range("C1") = "BOM component"
    ws.Range("D1") = "Component description"
    ws.Range("E1") = "Quantity Per 1000kg"
    ws.Range("F1") = "Quantity per 36,000 kg"
    ws.Range("G1") = "Component unit of measure"
    ws.Range("H1") = "Path"
    ws.Range("I1") = "Material Handling Type"
    ws.Range("J1") = "FP Component Type"
    ws.Range("K1") = "BP Origin Code"
End Sub

Sub getFormattedMaterialCode(ws)
    ws.Range("D1").EntireColumn.Insert
    ws.Range("D1") = "Formatted BOM Code"

    Dim endRow As Integer
    endRow = ws.Range("C2").End(xlDown).Row
    ws.Range("D2:D" & endRow).NumberFormat = "General"
    ws.Range("D2:D" & endRow).Formula = "=TEXT(C2, ""00000"")"
End Sub

Sub getWetBOMs(wbCompiledPath)
    Dim wbWetBOMPath As String
    wbWetBOMPath = wbCompiledPath & "/Wet BOMs/"
    readBOMs wbWetBOMPath, wbWetCompiled, "WP"
End Sub

Sub getDryBOMs(wbCompiledPath)
    Dim wbDryBOMPath As String
    wbDryBOMPath = wbCompiledPath & "/Dry BOMs/"
    readBOMs wbDryBOMPath, wbDryCompiled, "DB"
End Sub

Sub readBOMs(BOMPath, outputWS, powderType)
    Dim wbBOMName As String
    Dim wbBOMPath As String

    Dim nextRow As Integer
    wbBOMName = Dir(BOMPath & "*.xlsx")
    nextRow = 2
    Do While Len(wbBOMName) > 0
        If Right(wbBOMName, 4) = "xlsx" Then
            wbBOMPath = BOMPath + wbBOMName
            nextRow = grabBOMs(wbBOMPath, wbBOMName, outputWS, nextRow, powderType)
        End If
        wbBOMName = Dir
        DoEvents
    Loop
End Sub

Function grabBOMs(wbBOMPath, wbBOMName, outputWS, nextRow, powderType)
    Dim wbWorking As Workbook
    Dim ws As Worksheet

    Workbooks.Open FileName:=wbBOMPath, UpdateLinks:=False
    DoEvents
    Set wbWorking = Application.Workbooks(wbBOMName)
    For Each ws In wbWorking.Sheets
        nextRow = grabBOMsFromWS(ws, nextRow, outputWS, powderType)
    Next ws
    wbWorking.Close SaveChanges:=False
    DoEvents
    grabBOMs = nextRow
End Function

Function grabBOMsFromWS(BOMws, nextRow, outputWS, powderType)
    Dim ingredientRow As Integer
    Dim lastIngredientRow As Integer
    Dim cell As Range
    Dim colRange As Range
    Dim meteredColNumber As Integer
    Dim meteredColLetter As String

    Dim requiredInfo1 As Range
    Dim requiredInfo2 As Range
    Dim productName As String

    productName = BOMws.Name
    
    Dim lastColumnNumber As Integer
    Dim lastColumnLetter As String
    lastColumnNumber = BOMws.Range("A14").End(xlToRight).Column
    lastColumnLetter = Split(Cells(1, lastColumnNumber).Address, "$")(1)

    Set colRange = BOMws.Range("A14:" & lastColumnLetter & "14")
    meteredColNumber = Application.WorksheetFunction.Match("Material Handling Type", colRange, 0)
    meteredColLetter = Split(Cells(1, meteredColNumber).Address, "$")(1)

    Dim materialHandlingType As String
    lastIngredientRow = BOMws.Range("C14").End(xlDown).Row
    For Each cell In BOMws.Range("C15:C" & lastIngredientRow)

        If cell.Value = "*" Or cell.Font.Strikethrough = True Then
        Else
            ingredientRow = cell.Row
            materialHandlingType = BOMws.Range(meteredColLetter & ingredientRow).Value
            If materialHandlingType Like "Weigh & Dispense" Or materialHandlingType Like "Weighed & Dispensed" Then
                materialHandlingType = "Weigh & Dispense"
            End If

            If materialHandlingType = "Metered" Or materialHandlingType = "Weigh & Dispense" Then
                If BOMws.Range("F" & ingredientRow).Value <> 0 Then
                    Set requiredInfo1 = BOMws.Range("C" & ingredientRow & ":H" & ingredientRow)
                    Set requiredInfo2 = BOMws.Range(meteredColLetter & ingredientRow)
                    nextRow = pasteToSheet(requiredInfo1, requiredInfo2, nextRow, productName, outputWS, powderType)
                End If
            End If
        End If
    Next cell
    grabBOMsFromWS = nextRow
End Function

Function pasteToSheet(range1, range2, nextRow, productName, outputWS, powderType)
    outputWS.Range("A" & nextRow).Value = powderType
    outputWS.Range("B" & nextRow).Value = productName

    range1.Copy
    outputWS.Range("C" & nextRow & ":H" & nextRow).PasteSpecial xlPasteValues
    outputWS.Range("G" & nextRow).Value = LCase(outputWS.Range("G" & nextRow))
    range2.Copy
    outputWS.Range("I" & nextRow).PasteSpecial xlPasteValues
    pasteToSheet = nextRow + 1
End Function

Sub matchDryOnWet()
    wbWetCompiled.Range("C:C").NumberFormat = "@"
    getBPOriginOnWet

    wbDryCompiled.Range("C:C").NumberFormat = "@"
    getBPMatchOnDry
End Sub

Sub getBPMatchOnDry()
    Dim lastRow As Integer
    Dim powderCode As String
    Dim BOMComponentCode As String
    Dim BOMComponentCodeLength As Integer
    Dim BPComponentType As String
    Dim BPName As String
    Dim cell As Range

    Dim prevColB As String 
    Dim currColB As String 
    Dim sameAsPreviousRow As Boolean

    lastRow = wbDryCompiled.Range("C1").End(xlDown).Row
    For Each cell In wbDryCompiled.Range("C2:C" & lastRow)
        BOMComponentCode = wbDryCompiled.Range("C" & cell.Row).Value
        
        prevColB = wbDryCompiled.Range("B" & cell.Row - 1).Value
        currColB = wbDryCompiled.Range("B" & cell.Row).Value

        If prevColB = currColB Then 
            sameAsPreviousRow = True
        Else
            sameAsPreviousRow = False
        End If 

        If Len(BOMComponentCode) < 8 Then
            BPComponentType = "RawIngredient"
            If sameAsPreviousRow = False Then
                BPName = wbDryCompiled.Range("B" & cell.Row).Value
            End If
        ElseIf BOMComponentCode = "20033706" Then
            BPComponentType = "RawIngredient"
            If sameAsPreviousRow = False Then
                BPName = wbDryCompiled.Range("B" & cell.Row).Value
            End If
        Else
            BPComponentType = "BP"
            BPName = Left(wbDryCompiled.Range("D" & cell.Row).Value, 4)
            If Left(BPName, 1) = "S" Or Left(BPName, 1) = "Y" Then
                BPName = BPName
            Else
                BPName = wbDryCompiled.Range("B" & cell.Row).Value
            End If
        End If
        wbDryCompiled.Range("J" & cell.Row).Value = BPComponentType
        wbDryCompiled.Range("K" & cell.Row).Value = BPName
    Next cell
End Sub

Sub getBPOriginOnWet()
    Dim lastRow As Integer
    lastRow = wbWetCompiled.Range("C1").End(xlDown).Row

    wbWetCompiled.Range("J2:J" & lastRow).Value = "RawIngredient"
    wbWetCompiled.Range("K2:K" & lastRow).Value = wbWetCompiled.Range("B2:B" & lastRow).Value
End Sub

Sub joinWetOnDry()
    wbDryWetCompiled.Range("L1") = "Final FP Code"
    Dim lastRow As Integer, BPRow As Integer
    Dim BPCode As String, ingredientType As String, FPCode As String
    Dim ingredientRow As Range
    
    lastRow = wbDryCompiled.Range("A1").End(xlDown).Row
    BPRow = 2
    Dim cell As Range
    For Each cell In wbDryCompiled.Range("A2:A" & lastRow)
        ingredientType = wbDryCompiled.Range("J" & cell.Row)
        BPCode = wbDryCompiled.Range("K" & cell.Row)
        FPCode = wbDryCompiled.Range("B" & cell.Row)
        If ingredientType = "BP" Then
            wbDryWetCompiled.Range("L" & BPRow).Value = FPCode
            BPRow = getBPIngredients(BPCode, BPRow, FPCode, cell.Row)
        Else
            Set ingredientRow = wbDryCompiled.Range("A" & cell.Row & ":K" & cell.Row)
            wbDryWetCompiled.Range("A" & BPRow & ":K" & BPRow).Value = ingredientRow.Value
            wbDryWetCompiled.Range("L" & BPRow).Value = FPCode
            BPRow = BPRow + 1
        End If
    Next cell
    wbDryWetCompiled.Range("C:C").NumberFormat = "@"
End Sub

Function getBPIngredients(BPCode, BPRow, FPCode, scaleRow)
    Dim BPCodeList As Range
    Set BPCodeList = wbWetCompiled.Range("B1:B" & wbWetCompiled.Range("B1").End(xlDown).Row)

    Dim firstBPCodeRow As Integer, lastBPCodeRow As Integer
    firstBPCodeRow = Application.WorksheetFunction.Match(BPCode, BPCodeList, 0)
    lastBPCodeRow = firstBPCodeRow
    Do While True
        If wbWetCompiled.Range("B" & lastBPCodeRow).Value <> BPCode Then
            lastBPCodeRow = lastBPCodeRow - 1
            Exit Do
        Else
            lastBPCodeRow = lastBPCodeRow + 1
        End If
    Loop

    Dim ingredientRow As Range
    Set ingredientRow = wbWetCompiled.Range("A" & firstBPCodeRow & ":K" & lastBPCodeRow)
    wbDryWetCompiled.Range("A" & BPRow & ":K" & BPRow + (lastBPCodeRow - firstBPCodeRow)).Value = ingredientRow.Value
    
    Dim scale1000 As Double, scale36000 As Double
    scale1000 = wbDryCompiled.Range("E" & scaleRow).Value / 1000
    scale36000 = wbDryCompiled.Range("F" & scaleRow).Value / 36000

    Dim cell As Range
    For Each cell In wbDryWetCompiled.Range("E" & BPRow & ":E" & BPRow + (lastBPCodeRow - firstBPCodeRow))
        cell.Value = cell.Value * scale1000
        wbDryWetCompiled.Range("F" & cell.Row).Value = wbDryWetCompiled.Range("F" & cell.Row).Value * scale36000
    Next cell

    wbDryWetCompiled.Range("L" & BPRow & ":L" & BPRow + (lastBPCodeRow - firstBPCodeRow)).Value = FPCode
    getBPIngredients = wbDryWetCompiled.Range("A1").End(xlDown).Row + 1
End Function
