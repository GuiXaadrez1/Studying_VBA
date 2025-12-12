' ====================================================================================================
' CONFIGURAÇÃO PARTE CRUZAMENTO NCM
' ====================================================================================================

Const SourceSheetName As String = "Itens das NF-es Recebidas - Aut" ' Planilha com os NCMs de origem
Const SourceColumn As Long = 7 ' Coluna G: Coluna com os NCMs a serem lidos (Origem)
Const OutputColumn As Long = 13 ' Coluna M: Coluna onde o resultado da redução será inserido (Saída)
Const StartRowSource As Long = 4 ' Linha de início da iteração na planilha de origem

Const ReductionSheetName As String = "ReducaoNCM" ' Planilha com as taxas de redução
Const ReductionCodeColumn As Long = 1 ' Coluna A: Coluna contendo os NCMs (Redução)
Const ReductionTaxColumn As Long = 7 ' Coluna G: Coluna contendo a taxa de redução (Redução)
Const StartRowReduction As Long = 2 ' Linha de início da iteração na planilha de redução


' ===========================================================
' FUNÇÕES AUXILIARES
' ===========================================================

' Padroniza NCM, removendo pontos e espaços para garantir compatibilidade com a collection de redução.
Private Function NormalizarNCM(ByVal codigo As String) As String
    Dim clean As String
    clean = Replace(Replace(Trim(codigo), ".", ""), " ", "")
    
    ' Usa RegEx para remover todos os caracteres que não são dígitos (mais eficiente)
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        With re
            .Pattern = "\D" ' Qualquer não-dígito
            .Global = True
        End With
    End If
    
    NormalizarNCM = re.Replace(clean, "")
End Function

' Gera os níveis de NCM para busca: 8, 7, 6, 5, 4, 2 dígitos.
Private Function GerarNiveisNCM(ByVal codigoNCM As String) As collection
    
    Dim clean As String
    Dim col As New collection
    Set GerarNiveisNCM = col
    
    clean = NormalizarNCM(codigoNCM)
    
    ' Adiciona os níveis do mais específico para o mais genérico
    If Len(clean) >= 8 Then col.Add Left$(clean, 8)
    If Len(clean) >= 7 Then col.Add Left$(clean, 7)
    If Len(clean) >= 6 Then col.Add Left$(clean, 6)
    If Len(clean) >= 5 Then col.Add Left$(clean, 5)
    If Len(clean) >= 4 Then col.Add Left$(clean, 4)
    If Len(clean) >= 2 Then col.Add Left$(clean, 2)
    
End Function

' Verifica se uma chave existe em uma Collection usando tratamento de erro.
Private Function ExisteChave(collectionData As collection, chave As String) As Boolean

    On Error GoTo ErrHandler
        
        ' Tenta acessar o item, se falhar, o erro será capturado
        Dim tmp As Variant
        tmp = collectionData(chave)
        
        ExisteChave = True
    
    Exit Function

ErrHandler:
    ExisteChave = False

End Function

' Cria a Collection da planilha de Redução, usando o NCM como chave e o valor de redução como Item.
Private Function BuildReductionCollection( _
    Optional ByVal Sheet As Worksheet = Nothing, _
    Optional StartCellIndex As Long = StartRowReduction, _
    Optional ColumnCode As Long = ReductionCodeColumn, _
    Optional ColumnTax As Long = ReductionTaxColumn _
) As collection
    
    Dim lastRow As Long
    Dim RangeToMatriz As Variant
    Dim iRows As Long
    Dim codigo As String
    Dim valorReducao As Variant

    If Sheet Is Nothing Then Set Sheet = Worksheets(ReductionSheetName)

    Set BuildReductionCollection = New collection
    
    lastRow = Sheet.Cells(Sheet.rows.Count, ColumnCode).End(xlUp).Row
    
    ' Carrega a área de dados em uma matriz para velocidade
    RangeToMatriz = Sheet.Range( _
                        Sheet.Cells(StartCellIndex, ColumnCode), _
                        Sheet.Cells(lastRow, ColumnTax) _
                      ).Value
    
    ' A matriz RangeToMatriz é 1-baseada e tem UBound(matriz, 2) colunas
    Dim codeColIndex As Long
    Dim taxColIndex As Long
    
    ' Calcula os índices das colunas dentro da matriz (diferença entre colunas do Excel)
    codeColIndex = 1
    taxColIndex = ColumnTax - ColumnCode + 1
    
    For iRows = 1 To UBound(RangeToMatriz, 1)
        
        codigo = NormalizarNCM(CStr(RangeToMatriz(iRows, codeColIndex)))
        valorReducao = RangeToMatriz(iRows, taxColIndex)
        
        If Len(codigo) > 0 Then
            On Error Resume Next
                ' Adiciona a redução. On Error evita que códigos duplicados causem erro.
                BuildReductionCollection.Add Item:=valorReducao, key:=codigo
            On Error GoTo 0
        End If
    Next iRows

End Function

' Função principal que lê NCMs, busca o nível de redução e retorna uma matriz de resultados.
Private Function CruzarNcmsPorNiveis( _
    Optional ByVal SheetSource As String = SourceSheetName, _
    Optional ByVal startRow As Long = StartRowSource, _
    Optional ByVal ColumnSource As Long = SourceColumn _
) As Variant
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim iRows As Long
    
    Dim matrizFonte As Variant
    Dim matrizSaida() As Variant
    
    Dim reductionCollection As collection
    Dim niveisCollection As collection
    
    Dim codNcm As String
    Dim nivel As Variant
    Dim resultado As Variant
    Dim found As Boolean
    
    Set ws = Worksheets(SheetSource)
    
    ' Última linha da coluna de origem (SourceColumn)
    lastRow = ws.Cells(ws.rows.Count, ColumnSource).End(xlUp).Row
    
    ' Carrega TODOS os valores da coluna de origem (NCMs) na matriz de entrada
    matrizFonte = ws.Range( _
                        ws.Cells(startRow, ColumnSource), _
                        ws.Cells(lastRow, ColumnSource) _
                    ).Value
    
    ' Cria a matriz que receberá o resultado (1 coluna)
    ReDim matrizSaida(1 To UBound(matrizFonte, 1), 1 To 1)
    
    ' Collection da planilha ReducaoNCM (carregada uma vez)
    Set reductionCollection = BuildReductionCollection(Worksheets(ReductionSheetName))
    
    ' Loop principal
    For iRows = 1 To UBound(matrizFonte, 1)
        
        codNcm = NormalizarNCM(CStr(matrizFonte(iRows, 1)))
        
        If Len(Trim(codNcm)) = 0 Then
            matrizSaida(iRows, 1) = "0%"
            GoTo Proximo
        End If
        
        ' Gera níveis para o NCM (8, 7, 6, 5, 4, 2)
        Set niveisCollection = GerarNiveisNCM(codNcm)
        
        found = False
        resultado = "0%"
        
        ' Testa nível por nível (do mais específico ao mais genérico)
        For Each nivel In niveisCollection
            
            If ExisteChave(reductionCollection, CStr(nivel)) Then
                resultado = reductionCollection(CStr(nivel))
                found = True
                Exit For ' Encontrou o mais específico, pode parar a busca
            End If
        Next nivel
        
        ' Formata o resultado para a saída
        If found Then
            ' Formata para porcentagem (resultado é um decimal, p.ex., 0.10)
            matrizSaida(iRows, 1) = Format(CDbl(resultado), "0%")
        Else
            matrizSaida(iRows, 1) = "0%"
        End If
                
Proximo:
    Next iRows
    
    CruzarNcmsPorNiveis = matrizSaida

End Function

Public Sub ExecutarCruzamentoNCM()
    
    Dim ws As Worksheet
    Dim resultado As Variant
    Dim startRow As Long
    Dim lastRow As Long
    
    ' Desliga otimizações para agilizar o processo
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    On Error GoTo ErrHandler
    
    Set ws = Worksheets(SourceSheetName)
    
    ' Chama a função de cruzamento que retorna a matriz de resultados
    resultado = CruzarNcmsPorNiveis()
    
    ' Descobre o intervalo de gravação
    lastRow = ws.Cells(ws.rows.Count, SourceColumn).End(xlUp).Row
    startRow = StartRowSource
    
    ' Grava a matriz completa de resultados (saída) direto na planilha de uma vez
    ws.Range( _
        ws.Cells(startRow, OutputColumn), _
        ws.Cells(lastRow, OutputColumn) _
    ).Value = resultado
    
CleanExit:
    ' Restaura as configurações originais do Excel
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .Interactive = True
    End With
    
    ws.Activate
    MsgBox "Cruzamento concluído com sucesso!", vbInformation
    Exit Sub
    
ErrHandler:
    MsgBox "Ocorreu um erro: " & Err.Description & vbCrLf & "Verifique se as planilhas e colunas '" & SourceSheetName & "' e '" & ReductionSheetName & "' estão corretas.", vbCritical
    Resume CleanExit

End Sub
