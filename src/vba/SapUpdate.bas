Attribute VB_Name = "SapUpdate"

Option Explicit
Sub SAP_ME5A_UPDATE()

    On Error GoTo trataErro
  
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim requisitantes As Variant
    Dim linhaDestino As Long
    Dim grid As Object
    Dim msg As String
    Dim wsRelatorio As Worksheet
    Dim rng As Range
    Dim rng2 As Range
    
    Dim SapGuiAuto As Object
    Dim NApplication As Object
    Dim NConnection As Object
    Dim session As Object
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set NApplication = SapGuiAuto.GetScriptingEngine
    Set NConnection = NApplication.Children(0)
    Set session = NConnection.Children(0)
    
    requisitantes = Array("USUARIO01", "USUARIO02", "USUARIO03")
    Set wsRelatorio = ThisWorkbook.Sheets("Relatorio")
    
    wsRelatorio.Range("A2:G120").ClearContents
    
    For i = 0 To UBound(requisitantes)
    
        'Pesquisando a transacao + Enter
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nME5A"
        session.findById("wnd[0]").sendVKey 0
        
        'Tela da transacao ME5A
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = ""
        session.findById("wnd[0]/usr/chkP_FREIG").Selected = True 'Marcar as caixas de selecao
        session.findById("wnd[0]/usr/chkP_MEMORY").Selected = True
        session.findById("wnd[0]/usr/txtP_AFNAM").Text = requisitantes(i) 'Requisitante, roda novamente selecionando o proximo da lista
        session.findById("wnd[0]").sendVKey 8
        
        'Se caso aparecer a msg/sbar, significa que nao tem RCs desse requisitante
        msg = session.findById("wnd[0]/sbar").Text
    
        If msg <> "" Then
            GoTo ProximoReq
        End If
        
        Set grid = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell")
        
        'Selecionando todo relatorio e copiando
        grid.SelectAll
        grid.contextMenu
        grid.selectContextMenuItemByPosition "0"

        'Colando na planilha
        linhaDestino = wsRelatorio.Cells(wsRelatorio.Rows.Count, "A").End(xlUp).Row + 1
        
        Application.Wait Now + TimeValue("00:00:01")
        
        wsRelatorio.Cells(linhaDestino, 1).PasteSpecial
                
ProximoReq:
    Next i
    
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    'Formatar o intervalo
    Set rng = wsRelatorio.Range("A2:G120")
    Set rng2 = wsRelatorio.Range("F2:G120")
    
    
    'Adicionar aqui o filtro de apenas RCs com codigo de material _
    Linha abaixo se refere ao filtro da coluna "Material" que tira os vazios:
    rng.AutoFilter Field:=2, Criteria1:="<>"
    
    'Selecionando Range para editar no proximo bloco de codigo
    With wsRelatorio
        .Range("A1").Select
        .Range(Selection, Selection.End(xlToRight)).Select
        .Range(Selection, Selection.End(xlDown)).Select
    End With
    
    'Aqui ta colocando as grades da planilha
    With rng.Borders
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    rng2.Select

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With
    
    '".Sort" (Classifcar) 
    '"Key1:=.Range("H1")" (Coluna a ser classificada) 
    '"Order1:=xlAscending" (Ordem Crescente) 
    '"Header:=xlYes" (Tem cabecalhos)
    
    With wsRelatorio
        .Range("A1:H120").Sort Key1:=.Range("H1"), Order1:=xlAscending, Header:=xlYes
    End With
    
    wsRelatorio.Range("A1").Select
    
    'MsgBox "Transacao concluida!", vbInformation, "Sucesso!"
    
GoTo sair
    
sair:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub
    
trataErro:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Erro: " & Err.Description, vbCritical, "Erro no Processo, favor verificar ou chamar programador local!"
    
End Sub
