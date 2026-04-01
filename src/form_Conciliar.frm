VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Conciliar 
   Caption         =   "ConciliaÃ§Ã£o"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9270.001
   OleObjectBlob   =   "form_Conciliar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Conciliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Set ws = Sheets("JORNADA")

    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    With ListPendencias
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "50;80;80"
    End With

    For i = 2 To lastRow

        If Trim(UCase(ws.Cells(i, 9).Value)) = "INCOMPLETO" Then
            
            ListPendencias.AddItem
            ListPendencias.List(ListPendencias.ListCount - 1, 0) = i
            ListPendencias.List(ListPendencias.ListCount - 1, 1) = ws.Cells(i, 1).Value
            ListPendencias.List(ListPendencias.ListCount - 1, 2) = ws.Cells(i, 2).Value

        End If

    Next i

End Sub

Private Sub ListPendencias_Click()

    If ListPendencias.ListIndex = -1 Then Exit Sub

    Dim linha As Long
    linha = ListPendencias.List(ListPendencias.ListIndex, 0)

    Dim ws As Worksheet
    Set ws = Sheets("JORNADA")

    txtMatricula.Value = Format(ws.Cells(linha, 2).Value, "00000")
    txtNome.Value = BuscarNome(txtMatricula.Value)

    txtB1.Value = ws.Cells(linha, 4).Text
    txtB2.Value = ws.Cells(linha, 5).Text
    txtB3.Value = ws.Cells(linha, 6).Text
    txtB4.Value = ws.Cells(linha, 7).Text

End Sub

Private Sub btnSalvar_Click()

    If ListPendencias.ListIndex = -1 Then
        MsgBox "Selecione um registro.", vbExclamation
        Exit Sub
    End If

    Dim linha As Long
    linha = ListPendencias.List(ListPendencias.ListIndex, 0)

    Dim ws As Worksheet
    Set ws = Sheets("JORNADA")

    If txtB1.Value = "" Or txtB2.Value = "" Then
        MsgBox "Entrada e saída são obrigatórias.", vbExclamation
        Exit Sub
    End If

    ws.Cells(linha, 4).Value = TimeValue(txtB1.Value)
    ws.Cells(linha, 5).Value = TimeValue(txtB2.Value)

    If txtB3.Value <> "" Then ws.Cells(linha, 6).Value = TimeValue(txtB3.Value)
    If txtB4.Value <> "" Then ws.Cells(linha, 7).Value = TimeValue(txtB4.Value)

    Dim total As Double
    total = (ws.Cells(linha, 5).Value - ws.Cells(linha, 4).Value)

    If txtB3.Value <> "" And txtB4.Value <> "" Then
        total = total + (ws.Cells(linha, 7).Value - ws.Cells(linha, 6).Value)
    End If

    ws.Cells(linha, 8).Value = total
    ws.Cells(linha, 9).Value = "OK"

    ListPendencias.RemoveItem ListPendencias.ListIndex

    LimparCampos

    MsgBox "Registro conciliado.", vbInformation
    
      
    ' LIMPEZA DE SEGURANÇA:
    ListPendencias.ListIndex = -1 ' Remove o destaque da linha salva
    txtMatricula.Value = "" ' Limpa os campos para a próxima seleção
    txtNome.Value = ""
    ' ... limpe os demais campos de hora ...
    
    DoEvents

End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub

Private Function BuscarNome(mat As String) As String

    Dim ws As Worksheet
    Set ws = Sheets("FUNCIONARIOS")

    Dim i As Long

    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Format(ws.Cells(i, 1).Value, "00000") = Format(mat, "00000") Then
            BuscarNome = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i

    BuscarNome = ""

End Function

Private Function FormatHora(valor As Variant) As String

    If IsDate(valor) Then
        FormatHora = Format(valor, "hh:mm")
    Else
        FormatHora = ""
    End If

End Function

Private Sub LimparCampos()

    txtMatricula.Value = ""
    txtNome.Value = ""
    txtB1.Value = ""
    txtB2.Value = ""
    txtB3.Value = ""
    txtB4.Value = ""

End Sub



Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Sheets("DASHBOARD").Activate

End Sub




