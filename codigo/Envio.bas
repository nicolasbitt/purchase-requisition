Attribute VB_Name = "Envio"

Option Explicit
Sub Email()

    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim OutlookApp As Object
    Dim Email As Object
    Dim hoje As Date
    Dim bomDT As String
    Dim semana As String
    
    hoje = Date
    semana = Format(Date, "ww")
    
    'Copiando o range depois dos filtros para colocar no corpo do e-mail
    Range("A1:I1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    If TimeValue(Now) <= TimeValue("12:30") Then
        bomDT = "Bom dia, "
    Else
        bomDT = "Boa tarde, "
    End If
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set Email = OutlookApp.CreateItem(0)
    
      Email.To = "pessoa1@empresa.com;pessoa2@empresa.com;pessoa3@empresa.com;pessoa4@empresa.com;pessoa5@empresa.com;pessoa5@empresa.com"
      Email.CC = "grupo1@empresa.com;grupo2@empresa.com" 'Caso queira colocar alguem em copia inserir entre as aspas.
      Email.Subject = "Weekly Report ME5A - " & semana
  
      With Email
      .Display
  
        With .GetInspector.WordEditor.Application.Selection 'Acessa janela do Outlook -> Utiliza o WordEditor para alterar o corpo do e-mail -> WordEditor e um Documento que pertence a "Application" -> Seleciona o cursor do Word para depois colarmos no local correto.
          .TypeText bomDT & vbCrLf & vbCrLf
          .TypeText "Prezado comprador, favor verificar se existe alguma requisicao de compra pendente em sua carteira." & vbCrLf
          .TypeText "Caso haja alguma pendencia, favor dar seguimento ao processo de geracao do pedido. Caso contrario, desconsidere este e-mail." & vbCrLf & vbCrLf
          .Paste
          .TypeText vbCrLf & vbCrLf & "Qualquer duvida, estamos a disposicao." & vbCrLf & "Obrigado!" 'Se quiser colocar assinatura adicional no final
        End With
    End With
      
    Application.CutCopyMode = False
      
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub
