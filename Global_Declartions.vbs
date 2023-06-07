Attribute VB_Name = "Global_Declartions"
Public count As Integer

Public selectedFile As Workbook

Public Property Get mainWb() As Workbook

    Set mainWb = ThisWorkbook

End Property

Public Property Get mainWb_Report() As Worksheet

    Set mainWb_Report = mainWb.Sheets("Report")

End Property

Public Property Get downloadWb() As Workbook
    
    On Error Resume Next
    Set downloadWb = Workbooks.Open(mainWb.Path & "\Depreciation Simulation Download.xlsx")
    

    If Err.Number = "1004" Then
    
        count = count + 1
        
        If count < 2 Then
        
            MsgBox "Unable to find Depreciation Simulation Download file in the same folder (Check the Naming - 'Depreciation Simulation Download.xlsx'). Please select the file.", vbInformation, "File not Found"
            
            Dim fd As FileDialog
            Set fd = Application.FileDialog(msoFileDialogFilePicker)
            With fd
                .AllowMultiSelect = False
                .Filters.Add "Excel files", "*.xlsx"
            End With
            
            If fd.Show = -1 Then
                
                Set selectedFile = Workbooks.Open(fd.SelectedItems(1))
                Set downloadWb = selectedFile
                
            Else
            
                MsgBox "Download file not selected", vbOKOnly, "No file selected"
                
                End
                
            End If
            
        Else
        
            Set downloadWb = selectedFile
            
        End If
        
    End If

End Property

Public Property Get PD() As String

    PD = Trim(Mid(mainWb.Name, Application.WorksheetFunction.Find("PD", mainWb.Name, 1), 4))

End Property

Public Property Get FY() As String

    FY = Trim(Mid(mainWb.Name, Application.WorksheetFunction.Find("FY", mainWb.Name, 1), 4))

End Property

Public Function lastRow(ws As Worksheet, Col As Long) As Long

    lastRow = ws.Cells(Rows.count, Col).End(xlUp).row

End Function

Public Function lastCol(ws As Worksheet, r As Long) As Long

    lastCol = ws.Cells(r, Columns.count).End(xlToLeft).Column

End Function

Public Property Get lastRowDownloadWb() As Long

    lastRowDownloadWb = lastRow(downloadWb.Sheets(1), 1)
    
End Property

Public Property Get lastRowMainWb() As Long

    lastRowMainWb = lastRow(mainWb_Report, 1)
    
End Property

Public Property Get lastColDownloadWb() As Long

    lastColDownloadWb = lastCol(downloadWb.Sheets(1), 1)
    
End Property

Public Property Get APCcol() As Long

    APCcol = downloadWb.Sheets(1).Range(Cells(1, 1), Cells(1, lastColDownloadWb)).Find(what:="Cum. APC/RV").Column

End Property

Public Property Get AcqPdCol() As Long

    AcqPdCol = downloadWb_Sheet.Range(Cells(1, 1), Cells(1, lastColDownloadWb)).Find(what:="Acquis. " & Right(PD, 2) & "/" & Right(FY, 2)).Column
    
End Property

Public Property Get downloadWb_Sheet() As Worksheet

    Set downloadWb_Sheet = downloadWb.Sheets(1)

End Property

Public Property Get coCodeCol_DownloadWb() As Long

    coCodeCol_DownloadWb = downloadWb_Sheet.Range("A1:AZ1").Find(what:="Company Code", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get assetCol_DownloadWb() As Long

    assetCol_DownloadWb = downloadWb_Sheet.Range("A1:AZ1").Find(what:="Asset", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get assetDescCol_DownloadWb() As Long

    assetDescCol_DownloadWb = downloadWb_Sheet.Range("A1:AZ1").Find(what:="Asset Description", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get coCodeCol_MainWb() As Long

    coCodeCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Company Code", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get assetCol_MainWb() As Long

    assetCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Asset", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get assetDescCol_MainWb() As Long

    assetDescCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Asset Description", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get acqCostCol_MainWb() As Long

    acqCostCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Acquisition cost", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get accDepCol_MainWb() As Long

    accDepCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Accumulated Depreciation", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get NbvCol_MainWb() As Long

    NbvCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="NBV", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get curPdDepCol_MainWb() As Long

    curPdDepCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="Dep ", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get remarkCol_MainWb() As Long

    remarkCol_MainWb = mainWb_Report.Range("A1:AZ1").Find(what:="REMARKS", LookIn:=xlValues, lookat:=xlPart).Column
    
End Property

Public Property Get accDepCol_downloadWb() As Long

    accDepCol_downloadWb = downloadWb_Sheet.Range("A1:AZ1").Find(what:="Acc.dep.", LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get curPdDepCol_downloadWb() As Long

    curPdDepCol_downloadWb = downloadWb_Sheet.Range("A1:AZ1").Find(what:="Dep " & Right(PD, 2) & "/" & Right(FY, 2), LookIn:=xlValues, lookat:=xlWhole).Column

End Property

Public Property Get filteredData() As Range

    Set filteredData = mainWb_Report.Range("A2:H" & lastRowMainWb).Rows.SpecialCells(xlCellTypeVisible)
    If filteredData.Cells.count <= 8 Then
        Set filteredData = Nothing
    End If

End Property

Public Property Get filteredRemarkRng() As Range

    Set filteredRemarkRng = mainWb_Report.Range("H2:H" & lastRowMainWb).Rows.SpecialCells(xlCellTypeVisible)

End Property

Public Function Filter_Off()

    mainWb_Report.AutoFilterMode = False

End Function


