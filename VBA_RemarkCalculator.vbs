Attribute VB_Name = "VBA_RemarkCalculator"
Sub RemarkCalculator()


   If mainWb_Report.AutoFilterMode Then
    
        mainWb_Report.AutoFilterMode = False
        
   End If
   
   'AUC asset - Dep not required
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=assetCol_MainWb, Criteria1:=">80000000"
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:=">0"
    End With
    
    If Not filteredData Is Nothing Then
    
        filteredRemarkRng.Value = "AUC ASSET - DEP NOT REQUIRED"
                
    End If
    Filter_Off
        
    
    'AUC asset with asset cost 0 , hence to be removed
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=assetCol_MainWb, Criteria1:=">80000000"
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:="0"
    End With
    
    If Not filteredData Is Nothing Then
    
        filteredData.EntireRow.Delete
                
    End If
    Filter_Off
        
    'Assets fully depreciated
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=NbvCol_MainWb, Criteria1:="0"
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:=">0"
        .AutoFilter field:=curPdDepCol_MainWb, Criteria1:="0"
    End With
        
    If Not filteredData Is Nothing Then
        
        filteredRemarkRng = "ASSET COMPLETELY DEPRECIATED - OK"
    
    End If
    Filter_Off
        
        
    'Assets fully disposed - remove these lines
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:="0"
        .AutoFilter field:=accDepCol_MainWb, Criteria1:="<0"
        .AutoFilter field:=NbvCol_MainWb, Criteria1:="<0"
    End With
    If Not filteredData Is Nothing Then
        
        filteredData.EntireRow.Delete
        
    End If
    Filter_Off
        
        
    'Dummy asset cells to be removed
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:="0"
        .AutoFilter field:=accDepCol_MainWb, Criteria1:="0"
        .AutoFilter field:=NbvCol_MainWb, Criteria1:="0"
    End With
    
    If Not filteredData Is Nothing Then
        
        filteredData.EntireRow.Delete
        
    End If
    Filter_Off
    
        
    'Assets full depreciated but partically retired
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=curPdDepCol_MainWb, Criteria1:="0"
        .AutoFilter field:=NbvCol_MainWb, Criteria1:="<0"
    End With
    
    If Not filteredData Is Nothing Then
        
        filteredRemarkRng.Value = "ASSET COMPLETELY DEPRECIATED - OK - PARTIALLY RETIRED"
        
    End If
    Filter_Off
        
        
    'Asset partially disposed , depreciation is planned for current PD
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=curPdDepCol_MainWb, Criteria1:="<0"
        .AutoFilter field:=NbvCol_MainWb, Criteria1:="<0"
    End With
    If Not filteredData Is Nothing Then
        
        filteredRemarkRng.Value = "DEPRECIATION PLANNED POSTING IS AVAILABLE IN " & PD & " - OK - PARTIALLY DISPOSED"
        
    End If
    Filter_Off
    
    
    'New adds post biweekly depeciation run
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=curPdDepCol_MainWb, Criteria1:="0"
        .AutoFilter field:=accDepCol_MainWb, Criteria1:="0"
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:=">0"
    End With
    
    If Not filteredData Is Nothing Then
        
        If Application.WorksheetFunction.Sum(mainWb_Report.Range("D2:D" & lastRowMainWb).SpecialCells(xlCellTypeVisible)) _
        = Application.WorksheetFunction.Sum(mainWb_Report.Range("F2:F" & lastRowMainWb).SpecialCells(xlCellTypeVisible)) Then
            
            filteredRemarkRng.Value = "NEW ADDS AFTER 1102 BIWEEKLY DEP RUN"
            
        Else
        
            filteredRemarkRng.Value = "NEW ADDS AFTER 1102 BIWEEKLY DEP RUN - Sum Not Matching"
            
        End If
        
    End If
    Filter_Off
        
        
    'Assets planned for Depreciation run for current PD
    With mainWb_Report.Range("A1:H" & lastRowMainWb)
        .AutoFilter field:=remarkCol_MainWb, Criteria1:="="
        .AutoFilter field:=curPdDepCol_MainWb, Criteria1:="<0"
        .AutoFilter field:=accDepCol_MainWb, Criteria1:="<0", Operator:=xlOr, Criteria2:="0"
        .AutoFilter field:=acqCostCol_MainWb, Criteria1:=">0"
        .AutoFilter field:=NbvCol_MainWb, Criteria1:=">0", Operator:=xlOr, Criteria2:="0"
    End With
    If Not filteredData Is Nothing Then
        
        filteredRemarkRng.Value = "DEPRECIATION PLANNED POSTING IS AVAILABLE IN " & PD & "- OK"
        
    End If
    Filter_Off
     
     
End Sub

