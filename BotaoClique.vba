Sub BotaoClique_Click()
    Dim obj As Object
    
    Application.ScreenUpdating = False
    On Error GoTo ErroHandler
    
   
    Set obj = CreateObject("ProtecaoVBA.Calculador")
    obj.contaum Application
    
    Set obj = Nothing
    Application.ScreenUpdating = True
    Exit Sub

ErroHandler:
    Application.ScreenUpdating = True
    MsgBox "Erro ao executar automação: " & Err.Description, vbCritical
End Sub
