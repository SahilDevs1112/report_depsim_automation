Attribute VB_Name = "VBA_Formatting"
Sub MainWbFormatting_1()
    
    mainWb_Report.Cells(1, curPdDepCol_MainWb).Value = "Dep " & Right(PD, 2) & "/" & Right(FY, 2)
    
    If lastRow(mainWb_Report, 1) > 1 Then
    
        mainWb_Report.Range("A2:XFD" & mainWb_Report.Rows.count).Clear

    End If
    
End Sub

Sub deleteYellowLines()

    downloadWb_Sheet.Range("A1:AK" & lastRowDownloadWb).AutoFilter field:=1, Criteria1:=RGB(255, 255, 153), Operator:=xlFilterCellColor
    
    On Error Resume Next
    downloadWb_Sheet.Range("A2:AK" & lastRowDownloadWb + 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
    downloadWb_Sheet.AutoFilterMode = False
    
End Sub

Sub convertToNumber()

    With mainWb_Report.UsedRange
        .NumberFormat = "General"
        .Value = .Value
    End With

End Sub

Sub MainWbFormatting_2()
    
    If mainWb_Report.Cells(lastRowMainWb, 1).Value = "" Then
        mainWb_Report.Cells(lastRowMainWb, remarkCol_MainWb).Clear
    End If
    
    With mainWb_Report.UsedRange.Font
        .Name = "Calibri"
        .Size = 10
        .ThemeFont = xlThemeFontMinor
    End With
    ActiveWindow.Zoom = 80
    
    mainWb_Report.Activate
    mainWb_Report.Range("A2").Select
    ActiveWindow.FreezePanes = True
    
    mainWb_Report.Range("D2:G2").EntireColumn.Select
    Selection.Style = "Comma"
    
    mainWb_Report.UsedRange.Columns.AutoFit
    
    mainWb_Report.Range("A1:H1").Interior.ColorIndex = 15
    mainWb_Report.Range("D1:E" & lastRowMainWb).Interior.ColorIndex = 6
    mainWb_Report.Range("F1:F" & lastRowMainWb).Interior.ColorIndex = 43
    
    
    With mainWb_Report.UsedRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

End Sub
