Sub BotaoClique_Click()
    Dim obj As Object
    
    ' Desliga atualização de tela para evitar travamentos
    Application.ScreenUpdating = False
    
    On Error GoTo ErroHandler
    
    ' Chama a classe da DLL registrada
    Set obj = CreateObject("ProtecaoVBA.Calculador")
    obj.ContaUm Application
    
    ' Finaliza
    Set obj = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErroHandler:
    Application.ScreenUpdating = True
    MsgBox "Erro ao registrar/executar: " & Err.Description, vbCritical
End Sub
