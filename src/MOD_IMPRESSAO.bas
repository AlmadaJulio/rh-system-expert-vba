Attribute VB_Name = "MOD_IMPRESSAO"
Option Explicit

' --- CONFIGURAÇÃO DE SENHA ÚNICA ---
Private Const MINHA_SENHA As String = "SUASENHA"

' Macro para Impressão Individual
Public Sub IMPRESSAO_INDIVIDUAL()
    Dim wsDash As Worksheet: Set wsDash = Sheets("DASHBOARD")
    Dim wsJ As Worksheet: Set wsJ = Sheets("JORNADA")
    Dim wsPDF As Worksheet: Set wsPDF = Sheets("MODELO_PDF")
    
    Dim mat As String: mat = Format(wsDash.Range("W4").Value, "00000")
    Dim mes As Integer: mes = wsDash.Range("W3").Value
    Dim ano As Integer: ano = wsDash.Range("W2").Value
    
    ' 1. TRAVA DE SEGURANÇA (Sua lógica original)
    Dim i As Long, ultLinha As Long
    ultLinha = wsJ.Cells(wsJ.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultLinha
        If Format(wsJ.Cells(i, 2).Value, "00000") = mat And _
           wsJ.Cells(i, 12).Value = mes And wsJ.Cells(i, 13).Value = ano Then
            If wsJ.Cells(i, 9).Value = "INCOMPLETO" Then
                MsgBox "Não é possível gerar o PDF! Existem dias 'INCOMPLETOS'.", vbCritical
                Exit Sub
            End If
        End If
    Next i

    ' 2. MOLDURA DE SEGURANÇA (Abertura)
    On Error GoTo TratarErro
    ThisWorkbook.Unprotect Password:=MINHA_SENHA
    wsPDF.Visible = xlSheetVisible ' Torna visível para o comando Export funcionar

    ' 3. EXECUÇÃO
    Preencher_Modelo_Base mat
    
    Dim caminhoTemp As String: caminhoTemp = Environ("Temp") & "\Ponto_Individual.pdf"
    wsPDF.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoTemp, OpenAfterPublish:=True

TratarErro:
    ' 4. MOLDURA DE SEGURANÇA (Fechamento)
    wsPDF.Visible = xlSheetVeryHidden
    ThisWorkbook.Protect Password:=MINHA_SENHA, Structure:=True
    If Err.Number <> 0 Then MsgBox "Erro: " & Err.Description, vbExclamation
End Sub

' Macro para Impressão em Lote
Public Sub IMPRESSAO_LOTE()
    Dim wsFunc As Worksheet: Set wsFunc = Sheets("FUNCIONARIOS")
    Dim wsPDF As Worksheet: Set wsPDF = Sheets("MODELO_PDF")
    Dim i As Long, ultLinha As Long, contador As Integer: contador = 0
    Dim listaPlanilhas() As String
    Dim caminhoTemporario As String: caminhoTemporario = Environ("Temp") & "\Lote_Conferencia.pdf"
    
    If MsgBox("Gerar arquivo único de todos os ativos?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    On Error GoTo ErroLote
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- MOLDURA DE SEGURANÇA (Abertura) ---
    ThisWorkbook.Unprotect Password:=MINHA_SENHA
    wsPDF.Visible = xlSheetVisible
    
    ultLinha = wsFunc.Cells(wsFunc.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To ultLinha
        If UCase(wsFunc.Cells(i, 5).Value) = "ATIVO" Then
            ' Chama sua sub-rotina de preenchimento
            Preencher_Modelo_Base Format(wsFunc.Cells(i, 1).Value, "00000")
            
            contador = contador + 1
            wsPDF.Copy After:=Sheets(Sheets.Count)
            ActiveSheet.Name = "Temp_" & contador
            ReDim Preserve listaPlanilhas(1 To contador)
            listaPlanilhas(contador) = ActiveSheet.Name
        End If
    Next i
    
    If contador > 0 Then
        Sheets(listaPlanilhas).Select
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminhoTemporario, OpenAfterPublish:=True
        Sheets(listaPlanilhas).Delete
    End If

ErroLote:
    ' --- MOLDURA DE SEGURANÇA (Fechamento) ---
    wsPDF.Visible = xlSheetVeryHidden
    ThisWorkbook.Protect Password:=MINHA_SENHA, Structure:=True
    Sheets("DASHBOARD").Select
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Err.Number = 0 Then MsgBox "Lote gerado!", vbInformation Else MsgBox "Erro: " & Err.Description
End Sub

' Sua Sub-rotina de preenchimento (Mantida a lógica original)
Private Sub Preencher_Modelo_Base(matricula As String)
    Dim wsJ As Worksheet: Set wsJ = Sheets("JORNADA")
    Dim wsAux As Worksheet: Set wsAux = Sheets("AUXILIAR")
    Dim wsPDF As Worksheet: Set wsPDF = Sheets("MODELO_PDF")
    Dim wsFunc As Worksheet: Set wsFunc = Sheets("FUNCIONARIOS")
    
    Dim i As Long, linhaPDF As Long, nome As String, cargo As String, unidade As String
    Dim mes As Integer: mes = wsAux.Range("B2").Value
    Dim ano As Integer: ano = wsAux.Range("B3").Value
    
    ' Limpa batidas (Aba já está desprotegida pela macro que chamou esta sub)
    wsPDF.Range("A8:G33").ClearContents
    
    ' Busca dados cadastrais
    For i = 2 To wsFunc.Cells(wsFunc.Rows.Count, 1).End(xlUp).Row
        If Format(wsFunc.Cells(i, 1).Value, "00000") = matricula Then
            nome = wsFunc.Cells(i, 2).Value
            cargo = wsFunc.Cells(i, 3).Value
            unidade = wsFunc.Cells(i, 4).Value
            Exit For
        End If
    Next i
    
    ' Aplica mapeamento
    wsPDF.Range("B3").Value = unidade
    wsPDF.Range("B4").Value = nome
    wsPDF.Range("G2").Value = mes & "/" & ano
    wsPDF.Range("G3").Value = matricula
    wsPDF.Range("G4").Value = cargo
    
    ' Lança batidas
    linhaPDF = 8
    For i = 2 To wsJ.Cells(wsJ.Rows.Count, 1).End(xlUp).Row
        If Format(wsJ.Cells(i, 2).Value, "00000") = matricula And _
           wsJ.Cells(i, 12).Value = mes And wsJ.Cells(i, 13).Value = ano Then
           
            wsPDF.Cells(linhaPDF, 1).Value = Day(wsJ.Cells(i, 1).Value)
            wsPDF.Cells(linhaPDF, 2).Value = wsJ.Cells(i, 1).Value
            wsPDF.Cells(linhaPDF, 3).Value = wsJ.Cells(i, 4).Value
            wsPDF.Cells(linhaPDF, 4).Value = wsJ.Cells(i, 5).Value
            wsPDF.Cells(linhaPDF, 5).Value = wsJ.Cells(i, 6).Value
            wsPDF.Cells(linhaPDF, 6).Value = wsJ.Cells(i, 7).Value
            wsPDF.Cells(linhaPDF, 7).Value = wsJ.Cells(i, 8).Value
            linhaPDF = linhaPDF + 1
        End If
    Next i
End Sub

