VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_Cadastro 
   Caption         =   "Cadastro"
   ClientHeight    =   6600
   ClientLeft      =   180
   ClientTop       =   465
   ClientWidth     =   7605
   OleObjectBlob   =   "form_Cadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form_Cadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSalvar_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim novaLinha As ListRow
    
    ' Define a planilha e a tabela pelo nome que você deu
    Set ws = Sheets("FUNCIONARIOS")
    Set tbl = ws.ListObjects("tbl_Funcionarios")
    
    ' Validação básica
    If txtMatricula.Value = "" Or txtNome.Value = "" Then
        MsgBox "Preencha os campos obrigatórios.", vbExclamation
        Exit Sub
    End If
    
    ' O PULO DO GATO: Em vez de calcular "linha", adicionamos uma linha à tabela
    ' Se a tabela estiver vazia, ele preenche a primeira. Se tiver dados, ele cria a próxima.
    Set novaLinha = tbl.ListRows.Add
    
    ' Grava os dados usando a posição da coluna dentro da tabela
    With novaLinha
        .Range(1).Value = Format(txtMatricula.Value, "00000") ' Coluna 1
        .Range(2).Value = txtNome.Value                       ' Coluna 2
        .Range(3).Value = txtCargo.Value                      ' Coluna 3
        .Range(4).Value = txtUnidade.Value                    ' Coluna 4
        .Range(5).Value = cmbStatus.Value                     ' Coluna 5
    End With
    
    MsgBox "Cadastro realizado com sucesso!", vbInformation
    
    Unload Me
End Sub

Private Sub UserForm_Initialize()

    With cmbStatus
        .Clear
        .AddItem "ATIVO"
        .AddItem "INATIVO"
        .AddItem "AFASTADO"
    End With

End Sub
' Localiza o colaborador e sincroniza os campos internos da AUXILIAR
Public Sub Atualizar_Dados_Colaborador()
    Dim wsDash As Worksheet: Set wsDash = Sheets("DASHBOARD")
    Dim wsFunc As Worksheet: Set wsFunc = Sheets("FUNCIONARIOS")
    Dim wsAux As Worksheet: Set wsAux = Sheets("AUXILIAR")
    Dim matricula As String
    Dim i As Long, ultimaLinha As Long
    
    ' Captura a matrícula do Dashboard (W4) formatando com 5 dígitos
    matricula = Format(Val(wsDash.Range("W4").Value), "00000")
    
    ' Se W4 estiver vazio, limpa a F1 e encerra
    If wsDash.Range("W4").Value = "" Then
        wsAux.Range("F1").Value = "COLABORADOR: Selecione uma Matrícula"
        Exit Sub
    End If
    
    ' Busca na aba FUNCIONARIOS (Coluna A = Matrícula, Coluna B = Nome)
    ultimaLinha = wsFunc.Cells(wsFunc.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If Format(Val(wsFunc.Cells(i, 1).Value), "00000") = matricula Then
            ' 1. Escreve o Nome na F1 da AUXILIAR (para a caixa de texto ler)
            wsAux.Range("F1").Value = "COLABORADOR: " & wsFunc.Cells(i, 2).Value
            
            ' 2. Sincroniza os filtros internos da AUXILIAR (B1, B2, B3)
            wsAux.Range("B1").Value = matricula
            wsAux.Range("B2").Value = wsDash.Range("W3").Value ' Mês
            wsAux.Range("B3").Value = wsDash.Range("W2").Value ' Ano
            
            Exit Sub
        End If
    Next i
    
    ' Se percorrer tudo e não achar
    wsAux.Range("F1").Value = "COLABORADOR: Matrícula não encontrada"
End Sub

Private Sub btnSair_Click()

    Unload Me
    Sheets("DASHBOARD").Activate

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Sheets("DASHBOARD").Activate

End Sub


