Sub clearSheet()
    ActiveSheet.Range("A2", ActiveSheet.Range("G2").End(xlDown)).ClearContents
    ActiveSheet.EnableCalculation = True
End Sub

Sub CompileBlenderCampaign()

Application.ScreenUpdating = True
Application.DisplayAlerts = False
    
    currentrow = 0
    'iterate through each blender
    
    Sheets("Template").Range("A2:F3000").Clear
    For Each ws In Sheets
        If Left(ws.Name, 7) = "Blender" Then
            ws.Rows.EntireRow.Hidden = False
            ws.Columns.EntireColumn.Hidden = False
            'iterate through each campaign
            ncampaign = Application.WorksheetFunction.CountA(ws.Range("A2:A2000")) + 1
            For nrow = 2 To ncampaign
                
                BOM = ws.Cells(nrow, 2)
                bom_startime = ws.Cells(nrow, 5)
                bom_endtime = ws.Cells(nrow, 7)
                bom_totaltime = ws.Cells(nrow, 6)
                
                'for each campaign, add campaign number and its corresponding end time
                Sheets("Template").Cells(nrow + currentrow, 1) = Left(ws.Name, Len(ws.Name) - 8)
                Sheets("Template").Cells(nrow + currentrow, 2) = "Campaign " & ws.Range("A" & nrow)
                Sheets("Template").Cells(nrow + currentrow, 3) = ws.Cells(nrow, 2)
                Sheets("Template").Cells(nrow + currentrow, 4) = ws.Cells(nrow, 3)
                Sheets("Template").Cells(nrow + currentrow, 5) = ws.Cells(nrow, 4)
                Sheets("Template").Cells(nrow + currentrow, 6) = ws.Cells(nrow, 5)
                Sheets("Template").Cells(nrow + currentrow, 7) = ws.Cells(nrow, 7)
                Sheets("Template").Cells(nrow + currentrow, 8) = ws.Cells(nrow, 6)
            Next nrow
            currentrow = currentrow + ncampaign
        End If
    Next ws
    Sheets("Template").Select
    Sheets("Template").Columns("A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    Sheets("Template").Range("A2:H3000").Sort Key1:=Range("F2"), Order1:=xlAscending
    
    For Each material In Sheets("Material-Resource Mapping").Range("A2:A" & Sheets("Material-Resource Mapping").Range("A2").End(xlDown).Row)
        
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = CStr(material) Then
                Sheets(CStr(material)).Delete
            End If
        Next ws
        
        Sheets("Template").Copy after:=Sheets(Sheets.Count)
        Sheets("Template (2)").Name = CStr(material)
        nbom = Application.WorksheetFunction.CountA(Sheets(CStr(material)).Range("A2", Range("A2").End(xlDown)))
        For entry = 2 To nbom + 1
            nBOM_list = Application.WorksheetFunction.CountA(Sheets("BOM").Range("1:1"))
            colnum = 1
            For bom_check = 1 To nBOM_list Step 2
                If Sheets("BOM").Cells(1, colnum).Value <> Sheets(CStr(material)).Cells(entry, 3) Then
                    colnum = colnum + 2
                Else
                    bom_total = Application.Sum(Range(Sheets("BOM").Cells(2, colnum + 1), Sheets("BOM").Cells(2, colnum + 1).End(xlDown)))
                    For Item = 0 To 20
                    If CStr(Sheets("BOM").Cells(2 + Item, colnum)) = CStr(material) Or Sheets("BOM").Cells(2 + Item, colnum) = material Then
                            Usage = Sheets("BOM").Cells(2 + Item, colnum + 1)
                            Sheets(CStr(material)).Cells(entry, 9) = Usage * Int(Sheets(CStr(material)).Cells(entry, 4) * Sheets(CStr(material)).Cells(entry, 5) * 0.453592) / bom_total
                            Sheets(CStr(material)).Cells(entry, 10) = Sheets(CStr(material)).Cells(entry, 9) / Sheets(CStr(material)).Cells(entry, 8)
                            Exit For
                    End If
                    Next Item
                    colnum = 1
                End If
            Next bom_check
        
        If IsEmpty(Sheets(CStr(material)).Cells(entry, 9)) Then
            Sheets(CStr(material)).Cells(entry, 9) = 0
        End If
        If IsEmpty(Sheets(CStr(material)).Cells(entry, 10)) Then
            Sheets(CStr(material)).Cells(entry, 10) = 0
        End If
        
        If IsEmpty(Range("K2").Value) Then
            Sheets(CStr(material)).Cells(entry, 11).Formula = Sheets(CStr(material)).Cells(entry, 10)
        Else
            Sheets(CStr(material)).Cells(entry, 11).Formula = "=IF(F" & entry & ">" & "G" & entry - 1 & ",J" & entry & ",J" & entry & "+" & "J" & entry - 1 & ")"
        End If
        Sheets(CStr(material)).Cells(entry, 12).Formula = "=SUMIFS($K$2:K" & entry & ", $G$2:G" & entry & ", "">""&F" & entry & ",$F$2:F" & entry & ",""<""&G" & entry & ")"
        
        Next entry
        
    Next material
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
Sub Compile()

Sheets("Summary").Range("A2:M100000").Clear
For resource = 2 To Sheets("Material-Resource Mapping").Range("G2").End(xlDown).Row - 1
    emptyRow = Sheets("Summary").Range("B" & Rows.Count).End(xlUp).Offset(1).Row
    lastRow = Sheets("Template").Cells(Sheets("Template").Rows.Count, "A").End(xlUp).Row
    lastColumn = Sheets("Template").Cells(1, Sheets("Template").Columns.Count).End(xlToLeft).Column
    Sheets("Template").Range(Sheets("Template").Cells(2, 1), Sheets("Template").Cells(lastRow, lastColumn)).Copy
    Sheets("Summary").Range("B" & Rows.Count).End(xlUp).Offset(1).PasteSpecial xlPasteValues
    
    For j = emptyRow To lastRow + emptyRow - 2
        Sheets("Summary").Cells(j, 1) = Sheets("Material-Resource Mapping").Cells(resource, 7)
    Next j
    
    material = Split(Sheets("Material-Resource Mapping").Cells(resource, 8).Value, ",")
    For i = 0 To UBound(material)
        v = 2
        For r = emptyRow To emptyRow + lastRow - 2
            material(i) = LTrim(RTrim(material(i)))
            Sheets("Summary").Cells(r, 10).Value = Sheets("Summary").Cells(r, 10).Value + Sheets(CStr(material(i))).Cells(v, 9).Value
            Sheets("Summary").Cells(r, 11).Value = Sheets("Summary").Cells(r, 11).Value + Sheets(CStr(material(i))).Cells(v, 10).Value
            Sheets("Summary").Cells(r, 12).Value = Sheets("Summary").Cells(r, 12).Value + Sheets(CStr(material(i))).Cells(v, 11).Value
            Sheets("Summary").Cells(r, 13).Value = Sheets("Summary").Cells(r, 13).Value + Sheets(CStr(material(i))).Cells(v, 12).Value
            v = v + 1
        Next r
    Next i
    ' add resource
    'Exit Sub
Next
End Sub
