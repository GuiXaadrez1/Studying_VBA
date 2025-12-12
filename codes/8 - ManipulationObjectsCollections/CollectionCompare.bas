' ====================================================================================================
' CONFIGURAÇÃO PARTE CRUZAMENTO CNPJ
' ====================================================================================================

Const SourceSheetName As String = "Itens das NF-es Recebidas - Aut" ' Planilha que irevemos fazer o Cruzamento
Const SourceColumn As String = "G" ' Coluna G ao qual possui os Ncms a serem percorridos, comparados e cruzados
Const OutputColumn As String = "M" ' Coluna M ao qual vamos inserir as reduções conforme as comparações
Const StartRowSource As Long = 4 ' celula de inicio da iteracao

Const ReductionSheetName As String = "ReducaoNCM"
Const ReductionCodeColumn As String = "A"
Const ReductionTaxColumn As String = "G"
Const StartRowReduction As Long = 2

Const SheetCName As String = "PlanilhaC"
Const Ignore9DigitsInSheetC As Boolean = True


' ===========================================================
' FUNÇÕES AUXILIARES Para a Funcao CruzarNcm
' ===========================================================

Private Function SomenteDigitos(ByVal s As String) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    With re
        .Pattern = "\D"
        .Global = True
    End With
    SomenteDigitos = re.Replace(Trim$(s), "")
End Function

Private Function GerarNiveisNCM(ByVal codigo As String) As Collection
    Dim c As New Collection
    Dim t As Long: t = Len(codigo)
    
    If t >= 8 Then c.Add Left$(codigo, 8)
    If t >= 7 Then c.Add Left$(codigo, 7)
    If t >= 6 Then c.Add Left$(codigo, 6)
    If t >= 5 Then c.Add Left$(codigo, 5) ' gerenciar os genericos
    If t >= 4 Then c.Add Left$(codigo, 4)
    If t >= 2 Then c.Add Left$(codigo, 2)
    If t >= 1 Then c.Add Left$(codigo, 1)
    
    Set GerarNiveisNCM = c
End Function

Private Function ExisteChave(col As Collection, chave As String) As Boolean
    On Error GoTo ErrHandler
    Dim tmp As Variant
    tmp = col(chave)
    ExisteChave = True
    Exit Function
ErrHandler:
    ExisteChave = False
End Function

' Constroi a Collection de Reducao
Private Function BuildReductionCollection(ws As Worksheet) As Collection
    Dim col As New Collection
    Dim lastRow As Long
    lastRow = ws.Cells(ws.rows.Count, ReductionCodeColumn).End(xlUp).Row
    
    Dim i As Long, raw As String, norm As String, taxa As Variant
    Dim genKey As String
    
    For i = StartRowReduction To lastRow
        raw = CStr(ws.Cells(i, ReductionCodeColumn).Value)
        norm = SomenteDigitos(raw)
        
        If Len(norm) > 0 Then
            ' Ignorar planilha especial com 9 dígitos
            If ws.Name = SheetCName And Ignore9DigitsInSheetC And Len(norm) = 9 Then
                ' ignora
            Else
                taxa = ws.Cells(i, ReductionTaxColumn).Value
                
                ' Adiciona NCM completo
                If Not ExisteChave(col, norm) Then
                    col.Add Item:=taxa, key:=norm
                End If
                
                ' Adiciona genérico de 5 dígitos se tiver 5 ou mais
                If Len(norm) >= 5 Then
                    genKey = Left$(norm, 5)
                    If Not ExisteChave(col, genKey) Then
                        col.Add Item:=taxa, key:=genKey
                    End If
                End If
            End If
        End If
    Next i
    
    Set BuildReductionCollection = col
End Function

' ===========================================================
' FUNÇÕES PRINCIAPAL
' ===========================================================

Private Function CruzarNcmsPorNiveis(wsSrc As Worksheet, wsRed As Worksheet)

    On Error GoTo ErrHandler
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual ' tem a função de desativar o recálculo automático no Microsoft Excel
        Application.EnableEvents = False
        
        Dim colRed As Collection
        Set colRed = BuildReductionCollection(wsRed)
        
        Dim last As Long
        last = wsSrc.Cells(wsSrc.rows.Count, SourceColumn).End(xlUp).Row
        
        Dim i As Long, raw As String, norm As String
        Dim niveis As Collection, nivel As Variant
        Dim found As Boolean, result As Variant
        
        For i = StartRowSource To last
            raw = CStr(wsSrc.Cells(i, SourceColumn).Value)
            norm = SomenteDigitos(raw)
            
            If wsSrc.Name = SheetCName And Ignore9DigitsInSheetC And Len(norm) = 9 Then
                wsSrc.Cells(i, OutputColumn).Value = "Ignorado (9 dígitos)"
                GoTo NextRow
            End If
            
            If Len(norm) = 0 Then
                wsSrc.Cells(i, OutputColumn).Value = "0%"
                GoTo NextRow
            End If
            
            Set niveis = GerarNiveisNCM(norm)
            found = False
                        
            For Each nivel In niveis
                If Len(nivel) = 5 Then
                    ' Procurar qualquer NCM que comece com esse nível na Collection
                    Dim key As Variant
                    For Each key In colRed
                        If Left(key, 5) = CStr(nivel) Then
                            result = colRed(key)
                            found = True
                            Exit For
                        End If
                    Next key
                Else
                    If ExisteChave(colRed, CStr(nivel)) Then
                        result = colRed(CStr(nivel))
                        found = True
                    End If
                End If
                If found Then Exit For
            Next nivel
                        
            If found Then
                wsSrc.Cells(i, OutputColumn).Value = result
            Else
                wsSrc.Cells(i, OutputColumn).Value = "0%"
            End If
            
NextRow:
        Next i
        
        CruzarNcmComCollection = True
        
CleanExit:
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Exit Function
    
ErrHandler:
        MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical
        CruzarNcmComCollection = False
        Resume CleanExit

End Function

Public Sub ExecutarCruzamentoNCM()

    Dim wsSrc As Worksheet ' Planilha com a lista de NCMs para analisar
    Dim wsRed As Worksheet ' Planilha com a tabela de reduções
    Dim ok As Boolean

    ' **************************************
    ' 1. Obter as planilhas a partir das constantes
    ' **************************************
    On Error Resume Next
    Set wsSrc = ThisWorkbook.Worksheets(SourceSheetName)
    Set wsRed = ThisWorkbook.Worksheets(ReductionSheetName)
    On Error GoTo 0

    ' Validação de segurança
    If wsSrc Is Nothing Then
        MsgBox "A planilha de origem '" & SourceSheetName & "' não foi encontrada.", vbCritical
        Exit Sub
    End If

    If wsRed Is Nothing Then
        MsgBox "A planilha de redução '" & ReductionSheetName & "' não foi encontrada.", vbCritical
        Exit Sub
    End If

    ' **************************************
    ' 2. Limpar a coluna de saída
    ' **************************************
    wsSrc.Range(wsSrc.Cells(StartRowSource, OutputColumn), _
                wsSrc.Cells(wsSrc.rows.Count, OutputColumn)).ClearContents

    ' **************************************
    ' 3. Executar o cruzamento
    ' **************************************
    ok = CruzarNcmsPorNiveis(wsSrc, wsRed)

    ' **************************************
    ' 4. Mostrar mensagem final
    ' **************************************
    If ok Then
        MsgBox "Cruzamento concluído com sucesso!", vbInformation
    Else
        MsgBox "Ocorreu um problema durante o cruzamento.", vbExclamation
    End If

End Sub
