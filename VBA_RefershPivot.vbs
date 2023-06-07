Attribute VB_Name = "VBA_RefershPivot"
Sub RefreshPivot()
    
    mainWb.Sheets("Summary").PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook.PivotCaches.Create _
        (SourceType:=xlDatabase, SourceData:=mainWb.Sheets("Report").Range("A1:H" & lastRowMainWb))
        
    mainWb.Sheets("Summary").PivotTables("PivotTable1").RefreshTable
    
    With mainWb.Sheets("Summary").Range("B2")
    
        .Value = "SUMMARY - " & PD & " " & FY
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.ThemeFont = xlThemeFontMinor
        .Font.Bold = True
        
    End With

End Sub
