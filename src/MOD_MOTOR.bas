Attribute VB_Name = "MOD_MOTOR"
Option Explicit
Private Const MINHA_SENHA As String = "SUASENHA" ' <--- Ajuste para sua senha real

' =========================================================
' 1. MACRO PRINCIPAL (A ÚNICA QUE LIGA/DESLIGA A TELA)
' =========================================================

Public Sub BTN_CARREGAR_IMPORTACAO()
    Dim fd As FileDialog: Dim caminho As String
    Dim wbOrigem As Workbook: Dim wsOrigem As Worksheet: Dim wsDestino As Worksheet
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Selecione o arquivo do Forms"
        .Filters.Clear: .Filters.Add "Excel", "*.xlsx"
        If .Show <> -1 Then Exit Sub
        caminho = .SelectedItems(1)
    End With
    
    ' --- BLOQUEIO TOTAL PARA O VÍDEO ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    On Error GoTo ErroTratado
    Set wsDestino = ThisWorkbook.Sheets("IMPORTACAO")
    
    ' Limpa a aba de importação
    wsDestino.Rows("2:" & wsDestino.Rows.Count).ClearContents
    
    ' Abre e copia os dados silenciosamente
    Set wbOrigem = Workbooks.Open(Filename:=caminho, ReadOnly:=True)
    Set wsOrigem = wbOrigem.Sheets(1)
    Dim ult As Long: ult = wsOrigem.Cells(wsOrigem.Rows.Count, 1).End(xlUp).Row
    
    If ult >= 2 Then
        wsOrigem.Range("A2:C" & ult).Copy
        wsDestino.Range("A2").PasteSpecial xlPasteValues
    End If
    wbOrigem.Close False
    
    ' Executa o processamento pesado na memória (Ultra Rápido)
    IMPORTAR_MEMORIA_VELOZ
    GERAR_JORNADA_LIMPA_SILENCIOSA ' Certifique-se que o nome da sua macro de jornada é este
    
    ' --- FINALIZAÇÃO E ATUALIZAÇÃO ---
    Application.Calculation = xlCalculationAutomatic
    Calculate ' Atualiza B2/B3 na Auxiliar imediatamente
    Sheets("DASHBOARD").Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "Sincronização concluída!", vbInformation
    Exit Sub

ErroTratado:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

Private Sub IMPORTAR_MEMORIA_VELOZ()
    Dim wsImp As Worksheet: Set wsImp = Sheets("IMPORTACAO")
    Dim wsBase As Worksheet: Set wsBase = Sheets("BASE_PONTO")
    Dim dadosOrigem As Variant, dadosDestino As Variant
    Dim i As Long, ult As Long
    
    ult = wsImp.Cells(wsImp.Rows.Count, 1).End(xlUp).Row
    If ult < 2 Then Exit Sub
    
    ' Carrega tudo para a memória de uma vez
    dadosOrigem = wsImp.Range("A2:C" & ult).Value
    ReDim dadosDestino(1 To UBound(dadosOrigem, 1), 1 To 4)
    
    ' Processa na memória (milhares de vezes mais rápido que na célula)
    For i = 1 To UBound(dadosOrigem, 1)
        dadosDestino(i, 1) = Int(dadosOrigem(i, 1)) ' Data
        dadosDestino(i, 2) = dadosOrigem(i, 1) - Int(dadosOrigem(i, 1)) ' Hora
        dadosDestino(i, 3) = Format(dadosOrigem(i, 2), "00000") ' Matrícula 5 dígitos
        dadosDestino(i, 4) = dadosOrigem(i, 3) ' Unidade
    Next i
    
    ' Descarrega tudo na planilha de uma só vez
    wsBase.Rows("2:" & wsBase.Rows.Count).ClearContents
    wsBase.Columns("C:C").NumberFormat = "@"
    wsBase.Range("A2").Resize(UBound(dadosDestino, 1), 4).Value = dadosDestino
    
    ' Formatação rápida em bloco
    wsBase.Range("A2:A" & ult).NumberFormat = "dd/mm/yyyy"
    wsBase.Range("B2:B" & ult).NumberFormat = "hh:mm:ss"
End Sub

' =========================================================
' 2. IMPORTAÇÃO COM TRATAMENTO DE MATRÍCULA '00000'
' =========================================================
Private Sub IMPORTAR_LIMPO_SILENCIOSO()
    Dim wsImp As Worksheet: Set wsImp = Sheets("IMPORTACAO")
    Dim wsBase As Worksheet: Set wsBase = Sheets("BASE_PONTO")
    Dim lastRow As Long, i As Long
    
    lastRow = wsImp.Cells(wsImp.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    wsBase.Rows("2:" & wsBase.Rows.Count).ClearContents
    wsBase.Columns("C:C").NumberFormat = "@" ' Força texto para matrícula

    For i = 2 To lastRow
        wsBase.Cells(i, 1).Value = Int(wsImp.Cells(i, 1).Value)
        wsBase.Cells(i, 2).Value = wsImp.Cells(i, 1).Value - Int(wsImp.Cells(i, 1).Value)
        wsBase.Cells(i, 3).Value = Format(wsImp.Cells(i, 2).Value, "00000") ' Matrícula 5 dígitos
        wsBase.Cells(i, 4).Value = wsImp.Cells(i, 3).Value ' Unidade
    Next i
    
    wsBase.Range("A2:A" & lastRow).NumberFormat = "dd/mm/yyyy"
    wsBase.Range("B2:B" & lastRow).NumberFormat = "hh:mm:ss"
End Sub

' =========================================================
' 3. GERAÇÃO DE JORNADA (SEM REATIVAR A TELA)
' =========================================================
Private Sub GERAR_JORNADA_LIMPA_SILENCIOSA()
    Dim wsB As Worksheet: Set wsB = Sheets("BASE_PONTO")
    Dim wsJ As Worksheet: Set wsJ = Sheets("JORNADA")
    Dim lastRow As Long, i As Long, linha As Long
    Dim dict As Object, k As Variant, col As Collection
    Dim dataVal As Date, mat As String, chave As String
    Dim totalDia As Double, arr() As Double, x As Long

    wsJ.Rows("2:" & wsJ.Rows.Count).ClearContents
    lastRow = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Ordenação Cronológica
    With wsB.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsB.Range("A2:A" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=wsB.Range("C2:C" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=wsB.Range("B2:B" & lastRow), Order:=xlAscending
        .SetRange wsB.Range("A1:D" & lastRow): .Header = xlYes: .Apply
    End With

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        dataVal = wsB.Cells(i, 1).Value
        mat = Format(wsB.Cells(i, 3).Value, "00000")
        chave = CLng(dataVal) & "|" & mat
        If Not dict.exists(chave) Then dict.Add chave, New Collection
        dict(chave).Add wsB.Cells(i, 2).Value
    Next i

    linha = 2
    For Each k In dict.keys
        dataVal = CDate(Split(k, "|")(0))
        mat = Split(k, "|")(1)
        Set col = dict(k)
        ReDim arr(1 To col.Count)
        For x = 1 To col.Count: arr(x) = col(x): Next x
        QuickSort arr, LBound(arr), UBound(arr)

        wsJ.Cells(linha, 1).Value = dataVal
        wsJ.Cells(linha, 2).Value = mat
        ' Busca unidade direto na linha atual para ser mais rápido
        wsJ.Cells(linha, 3).Value = BUSCAR_UNIDADE_RAPIDO(wsB, dataVal, mat)

        For x = 1 To Application.WorksheetFunction.Min(4, UBound(arr))
            wsJ.Cells(linha, 3 + x).Value = arr(x)
        Next x

        totalDia = CALCULAR_HORAS_SEGURAS(arr)
        wsJ.Cells(linha, 8).Value = totalDia
        wsJ.Cells(linha, 9).Value = IIf(UBound(arr) >= 4, "OK", "INCOMPLETO")
        
        If wsJ.Cells(linha, 9).Value = "OK" Then
            If totalDia > (8 / 24) Then
                wsJ.Cells(linha, 10).Value = totalDia - (8 / 24)
                wsJ.Cells(linha, 11).Value = 0
            Else
                wsJ.Cells(linha, 10).Value = 0
                wsJ.Cells(linha, 11).Value = (8 / 24) - totalDia
            End If
        End If
        wsJ.Cells(linha, 12).Value = Month(dataVal)
        wsJ.Cells(linha, 13).Value = Year(dataVal)
        linha = linha + 1
    Next k

    wsJ.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    wsJ.Columns("D:H").NumberFormat = "[hh]:mm:ss"
    wsJ.Columns("J:K").NumberFormat = "[hh]:mm:ss"
End Sub

' --- FUNÇÕES DE APOIO (MANTIDAS) ---
Public Sub QuickSort(arr() As Double, first As Long, last As Long)
    Dim low As Long, high As Long, mid As Double, temp As Double
    low = first: high = last: mid = arr((first + last) \ 2)
    Do While low <= high
        Do While arr(low) < mid: low = low + 1: Loop
        Do While arr(high) > mid: high = high - 1: Loop
        If low <= high Then
            temp = arr(low): arr(low) = arr(high): arr(high) = temp
            low = low + 1: high = high - 1
        End If
    Loop
    If first < high Then QuickSort arr, first, high
    If low < last Then QuickSort arr, low, last
End Sub

Public Function BUSCAR_UNIDADE_RAPIDO(ws As Worksheet, dataVal As Date, mat As String) As String
    Dim i As Long, lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        If Int(ws.Cells(i, 1).Value) = Int(dataVal) And Format(ws.Cells(i, 3).Value, "00000") = mat Then
            BUSCAR_UNIDADE_RAPIDO = ws.Cells(i, 4).Value: Exit Function
        End If
    Next i
End Function

Public Function CALCULAR_HORAS_SEGURAS(arr() As Double) As Double
    Dim total As Double
    If UBound(arr) >= 2 Then If arr(2) > arr(1) Then total = arr(2) - arr(1)
    If UBound(arr) >= 4 Then If arr(4) > arr(3) Then total = total + (arr(4) - arr(3))
    CALCULAR_HORAS_SEGURAS = total
End Function
