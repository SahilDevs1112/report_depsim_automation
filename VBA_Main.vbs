Attribute VB_Name = "VBA_Main"
Sub Main()

    Application.ScreenUpdating = False
    
    Call MainWbFormatting_1
    
    Call deleteYellowLines
    
    Call TransferDataToMainWb_CopyPasteMethod
    
    Call convertToNumber
    
    Call RemarkCalculator
    
    Call MainWbFormatting_2
    
    Call RefreshPivot
          
    downloadWb.Close False
        
    Application.ScreenUpdating = True
        
            
End Sub
