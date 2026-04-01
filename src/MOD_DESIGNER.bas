Attribute VB_Name = "MOD_DESIGNER"
Option Explicit

' 1. LIBERA O ARQUIVO PARA AJUSTES
Sub MODO_DESIGNER_ON()
    Dim senha As String: senha = "SUASENHA" ' <-- COLOQUE SUA SENHA AQUI
    Dim aba As Worksheet
    
    On Error Resume Next
    ThisWorkbook.Unprotect senha
    For Each aba In ThisWorkbook.Worksheets
        aba.Visible = xlSheetVisible
        aba.Unprotect senha
    Next aba
    MsgBox "Modo Edição Ativado!", vbInformation
End Sub

' 2. PREPARA PARA A GRAVAÇÃO (SEM ERRO AGORA)
Sub MODO_GRAVACAO_ON()
    Dim aba As Worksheet
    Dim senha As String: senha = "SUASENHA" ' <-- COLOQUE SUA SENHA AQUI
    
    On Error Resume Next
    ThisWorkbook.Unprotect senha
    
    For Each aba In ThisWorkbook.Worksheets
        If aba.Name = "DASHBOARD" Or aba.Name = "VIEW" Then
            aba.Visible = xlSheetVisible
        Else
            aba.Visible = xlSheetVeryHidden
        End If
    Next aba
    
    Sheets("DASHBOARD").Activate
    ThisWorkbook.Protect senha, Structure:=True
    MsgBox "Sistema Pronto para Gravação!", vbInformation
End Sub
