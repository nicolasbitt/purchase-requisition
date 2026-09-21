Attribute VB_Name = "Hidden"

Option Explicit
Sub hiddenInfo()

    Dim wsInfo As Worksheet
    Dim wsPlan As Worksheet
    
    Set wsInfo = ThisWorkbook.Sheets(2)
    Set wsPlan = ThisWorkbook.Sheets(3)
    
    wsInfo.Visible = xlSheetVeryHidden
    wsPlan.Visible = xlSheetVeryHidden

End Sub
