Attribute VB_Name = "MOD_BOTOES"
Option Explicit


Public Sub BTN_CADASTRO()
    Sheets("VIEW").Activate
    form_Cadastro.Show
End Sub

Public Sub BTN_CONCILIAR()
    Sheets("VIEW").Activate
    form_Conciliar.Show
End Sub

Public Sub BTN_IMPRESSAO()
    IMPRESSAO_INDIVIDUAL
End Sub

Public Sub BTN_GERAR_LOTE()
    IMPRESSAO_LOTE
End Sub

Public Sub BTN_LIMPAR_FILTROS()

    Dim wsDash As Worksheet
    Set wsDash = Sheets("DASHBOARD")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Limpa apenas valores (mantém validação)
    wsDash.Range("W2:W4").ClearContents
    wsDash.Range("E1").ClearContents
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
 
    MsgBox "Filtros limpos com sucesso!", vbInformation

End Sub

