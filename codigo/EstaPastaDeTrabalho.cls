VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
    
    'Tira filtro ao abrir a planilha
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    Set wb = ThisWorbook
    Set ws = wb.Sheets(1)
    Set rng = ws.Range("$A$1:$K$120")
    
    rng.AutoFilter Field:=4

End Sub
