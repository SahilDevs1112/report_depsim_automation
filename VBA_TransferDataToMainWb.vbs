Attribute VB_Name = "VBA_TransferDataToMainWb"

Sub TransferDataToMainWb_loopMenthod()
    
    'not using this method because of its excecution time
    Dim currow As Long
    For currow = 2 To lastRowDownloadWb
    
        mainWb_Report.Cells(currow, coCodeCol_MainWb).Value = downloadWb_Sheet.Cells(currow, coCodeCol_DownloadWb).Value
        mainWb_Report.Cells(currow, assetCol_MainWb).Value = downloadWb_Sheet.Cells(currow, assetCol_DownloadWb).Value
        mainWb_Report.Cells(currow, assetDescCol_MainWb).Value = downloadWb_Sheet.Cells(currow, assetDescCol_DownloadWb).Value
        mainWb_Report.Cells(currow, curPdDepCol_MainWb).Value = downloadWb_Sheet.Cells(currow, curPdDepCol_downloadWb).Value
        
        'Calculate acqisition cost , acc dep and NBV
        mainWb_Report.Cells(currow, acqCostCol_MainWb).Value = Application.WorksheetFunction.Sum(downloadWb_Sheet.Range(Cells(currow, APCcol), Cells(currow, AcqPdCol)))
        mainWb_Report.Cells(currow, accDepCol_MainWb).Value = Application.WorksheetFunction.Sum(downloadWb_Sheet.Range(Cells(currow, accDepCol_downloadWb), Cells(currow, curPdDepCol_downloadWb - 1)))
        mainWb_Report.Cells(currow, NbvCol_MainWb).Value = mainWb_Report.Cells(currow, acqCostCol_MainWb) - mainWb_Report.Cells(currow, accDepCol_downloadWb)
          
    Next currow

End Sub

Sub TransferDataToMainWb_CopyPasteMethod()

    'copy Company code column
    downloadWb_Sheet.Range(Cells(2, coCodeCol_DownloadWb), Cells(lastRowDownloadWb, coCodeCol_DownloadWb)).Copy
    mainWb_Report.Cells(2, coCodeCol_MainWb).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'copy Asset column
    downloadWb_Sheet.Range(Cells(2, assetCol_DownloadWb), Cells(lastRowDownloadWb, assetCol_DownloadWb)).Copy
    mainWb_Report.Cells(2, assetCol_MainWb).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'copy Asset Description column
    downloadWb_Sheet.Range(Cells(2, assetDescCol_DownloadWb), Cells(lastRowDownloadWb, assetDescCol_DownloadWb)).Copy
    mainWb_Report.Cells(2, assetDescCol_MainWb).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'copy Current PD depreciation column
    downloadWb_Sheet.Range(Cells(2, curPdDepCol_downloadWb), Cells(lastRowDownloadWb, curPdDepCol_downloadWb)).Copy
    mainWb_Report.Cells(2, curPdDepCol_MainWb).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
        
    'Sum of aqusition cost from DownloadWb to MainWb
    downloadWb_Sheet.Cells(1, AcqPdCol + 1).EntireColumn.Insert
    downloadWb_Sheet.Range(Cells(2, AcqPdCol + 1), Cells(lastRowDownloadWb, AcqPdCol + 1)).Formula2R1C1 = "=SUM(RC[-1]:RC[-" & AcqPdCol + 1 - APCcol & "])"
    downloadWb_Sheet.Range(Cells(2, AcqPdCol + 1), Cells(lastRowDownloadWb, AcqPdCol + 1)).Copy
    mainWb_Report.Cells(2, acqCostCol_MainWb).PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'Sum of Accumulated Depreciation from DownloadWb to MainWb
    downloadWb_Sheet.Cells(1, curPdDepCol_downloadWb).EntireColumn.Insert
    downloadWb_Sheet.Range(Cells(2, curPdDepCol_downloadWb - 1), Cells(lastRowDownloadWb, curPdDepCol_downloadWb - 1)).Formula2R1C1 = "=SUM(RC[-1]:RC[-" & curPdDepCol_downloadWb - 1 - accDepCol_downloadWb & "])"
    downloadWb_Sheet.Range(Cells(2, curPdDepCol_downloadWb - 1), Cells(lastRowDownloadWb, curPdDepCol_downloadWb - 1)).Copy
    mainWb_Report.Cells(2, accDepCol_MainWb).PasteSpecial xlPasteValues
    
    'Calculate NBV in mainWb
    mainWb_Report.Range("F2:F" & lastRowMainWb).Formula2 = "=SUM($E2:$D2)"
    Range("a1").Select
End Sub
