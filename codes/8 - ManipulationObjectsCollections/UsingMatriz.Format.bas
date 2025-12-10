
' Criar uma código VBA que formate a coluna G de NCM da planilha principal'
' E insira parte do codigo em cada coluna especifica -> capitulo, posisao, etc...
'
' Depois vá para planilha ReducaoNCM, faca a mesma formatacao da planilha
' Por fim faca o cruzamento de NCMs por niveis!

''''''''''''''''''''''''
' Configuracoes Gerais '
''''''''''''''''''''''''

' nome da nossa Planilha principal
Const nameSheetPlanOne As String = "Itens das NF-es Recebidas - Aut"

' Inicia o range na celula 4 -> Planilha Principal
Const indexCellFirstColumn As Long = 4 ' linha 4

Const indexColumnPlanOne As Long = 7 ' Column G

' nome da nossa Planilha de Reducao de produto conforme seu codigo Ncm
Const nameSheetPlanTwo As String = "ReducaoNCM"

' Inicia o range na celula 4 -> Planilha ReducaoNCM
Const indexCellFirstColumnReducaoNCM As Long = 2 ' linha 2

Const indexCellColumnReducaoNCM As Long = 1 ' Coluna A

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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


Private Function ReadFormatarNcmPlanOne( _
    Optional ByVal nameSheet As String = nameSheetPlanOne, _
    Optional indexStartRange As Long = indexCellFirstColumn, _
    Optional indexExecutionColumnPlanOne As Long = indexColumnPlanOne _
) As Variant ' Vai retonar uma Matriz 2D redimensioando para mais 5 colunas contendo todo o intervalo de dados formatados em memoria

    ' Planilha Principal
    Dim Sheet As Worksheet
    
    ' Variavael que vai representar a matriz criada atraves de um objeto Range!
    Dim RangeToMatriz As Variant
    
    ' Variavel que vai Representar a matriz de saida comforme a matriz criada de um objeto Range
    Dim MatrizOut As Variant
    
    ' Representa a celula especifica da nossa range
    Dim cell As Variant
    
    ' Representa o iterador pela linha na nossa Matriz
    Dim iRow As Long
    
    ' Representa o iterador pela coluna na nossa Matriz
    Dim iCol As Long
    
    ' quantidade de caracteres que o codigo ncm possui nesta range, intervalo de celulas
    Dim qtdCaracter As Long
    
    ' Definindo a ultima linha do nosso intervalo de celula
    Dim lastRow As Long
    
    ' Usando a nossa planilha
    Set Sheet = Worksheets(nameSheet)
    
    ' Ativando essa planilha
    Sheet.Activate
    
     MsgBox "Clique Ok para formatar a Planilha: Itens das NF-es Recebidas - Aut"
    
    ' puxando a nossa ultima Linha!
    lastRow = Sheet.Cells(Sheet.rows.Count, indexExecutionColumnPlanOne).End(xlUp).Row
    
    ' Criando Matriz atraves do Range com Intervalo dinamico conforme a configuracao
    RangeToMatriz = Sheet.Range( _
        Sheet.Cells(indexStartRange, indexExecutionColumnPlanOne), _
        Sheet.Cells(lastRow, indexExecutionColumnPlanOne) _
    ).Value
    
    ' Percorrendo a Matriz para Debugg
    
    ' For iRow = LBound(RangeToMatriz, 1) To UBound(RangeToMatriz, 1)
   
        ' cell = RangeToMatriz(iRow, 1) ' colocando manualmente a iteracao na primeira coluna
        ' Debug.Print "Acessando a celula na Linha: "; CStr(cell)
   
    ' Next iRow
    
    ' Redimensionando a Matriz pra conte 5 colunas adicionais

    ' Redim -> Redimensiona a nossa matriz
    ' LBound(matriz,coluna) -> Pega o limite inferior (primeiro index da matriz)
    ' Ubound(matriz,coluna) -> Pega o limite superior (ultimo index da matriz) -> neste caso representando a ultima linha
    ' Primeira parte do redimensionamento... Linhas -> Vetor Unidimensional
    ' Segudna parte do redimensionamento... Colunas -> Vetor Multidimenisonal

    
    ' Verificando se a Matriz criada esta vazia!
    If IsEmpty(RangeToMatriz) Then
        MsgBox "Nenhum dado encontrado para formatar."
        Exit Function
    End If
    
    ' Descobrir dimensões reais para a MatrizOut
    Dim rows As Long
    rows = UBound(RangeToMatriz, 1)
    
    ' Redimensiona a Matriz criada com o Range com  saída de 5 colunas
    ReDim MatrizOut(1 To rows, 1 To 5)
       
    ' MsgBox "Degubando!" ' Até aqui passou
       
    For iRow = 1 To rows
        
        ' Acessando a celula
        cell = RangeToMatriz(iRow, 1)
         
        ' Realizando tratamento de dados na cell, removendo pontos e espaços internos na esquerda e direita
        cell = Trim(Replace(cell, ".", ""))
        
        qtdCaracter = Len(cell)
        
        ' Usando Select Case para Fazer as Validações
                
        Select Case qtdCaracter
            Case Is = 8: cell = cell
            Case Is = 7: cell = String(8 - Len(cell), "0") & cell
            Case Is = 6: cell = String(8 - Len(cell), "0") & cell
            Case Is = 5: cell = String(8 - Len(cell), "0") & cell
            Case Is = 4: cell = String(8 - Len(cell), "0") & cell
            Case Is = 3: cell = String(8 - Len(cell), "0") & cell
            Case Is = 2: cell = String(8 - Len(cell), "0") & cell
            Case Is = 1: cell = String(8 - Len(cell), "0") & cell
            
            ' String(quantidade_vezes_que_o_segundo_parametro_repete,string_a_repetir)
            
            Case Else
               MsgBox "NCM nao identificado na celula: " & CStr(Sheet.Cells(iRow, indexExecutionColumnPlanOne))
        End Select
        
        ' Se menor que 8 digitos... colocar zero a esqueda uma vez
        ' Se menor que 7.. zero a esquerda duas vezes... por ai vai.. ate ser igual a 0
    
        'matrizNCM(i, 1) = Left$(cell, 2) ' Retorna uma string contendo o número de caracteres definido em Tamanho do lado esquerdo da String.
        MatrizOut(iRow, 1) = Left$(cell, 2)
        MatrizOut(iRow, 2) = Mid$(cell, 3, 2)
        MatrizOut(iRow, 3) = Mid$(cell, 5, 2)
        MatrizOut(iRow, 4) = Mid$(cell, 7, 1)
        MatrizOut(iRow, 5) = Mid$(cell, 8, 1)
        
        ' Mid$(String,posicao_inxes_extracao,tamanho_extracao) -> é usada para extrair uma sub-cadeia (substring) de caracteres de dentro de uma cadeia de caracteres (string) maior.
        
        ' Por Fim Reformatando o NCM com todos os dígitos e mascaras!
        
        cell = Left$(cell, 2) & "." & Mid$(cell, 3, 2) & "." & Mid$(cell, 5, 2) & "." & Mid$(cell, 7, 1) & "." & Mid$(cell, 8, 1)
        
        Debug.Print "Acessando a celula na Linha: " & CStr(iRow)
        
ProximaLinha:
    Next iRow
    
    ReadFormatarNcmPlanOne = MatrizOut
    
End Function



Private Function ValidarRangeToFormat(ByVal interval As Range) As Boolean
    
    Dim Matriz As Variant
    Dim i As Long
    Dim j As Long
    
    Matriz = interval.Value
    
    
    For i = LBound(Matriz, 1) To UBound(Matriz, 1)
        For j = LBound(Matriz, 2) To UBound(Matriz, 2)
            
            If Matriz(i, j) = "" Then
                ValidarRangeToFormat = False
                Exit Function
            End If
        
        Next j
    Next i
    
    ' Se não encontrou vazios -> válido
    ValidarRangeToFormat = True

End Function

Private Function ReadFormatarPlanilhaReducaoNcm( _
    Optional ByVal sheetName As String = nameSheetPlanTwo, _
    Optional indexStartRange As Long = indexCellFirstColumn, _
    Optional indexExecutionColumnPlanOne As Long = indexColumnPlanOne _
) As Variant
    
    Dim Sheet As Worksheet
        
    Dim RangeToMatriz As Variant
     
    Dim cel As Variant
    'Dim valueCel As String
     
    ' Definindo a ultima linha do nosso intervalo de celula
    Dim lastRow As Long
    
    ' Iteradores
    Dim iRows As Long
    Dim iColumn As Long
    
    Dim qtdCracteres As Long
    
    ' Ativando a nossa planilha ReducaoNCM
    
    Set Sheet = Worksheets(sheetName)
        
    Sheet.Activate
    
    MsgBox "Clique Ok para formatar a Planilha: ReducaoNCM"
                
    ' Definindo a ultima linha Preenchida
    lastRow = Sheet.Cells(Sheet.rows.Count, indexExecutionColumnReducaoNCM).End(xlUp).Row
    
    ' Definindo o Range/Intervalo da coluna conforme a configuracao
    RangeToMatriz = Sheet.Range(Sheet.Cells(indexStartRangeReducaoNCM, indexExecutionColumnReducaoNCM), Sheet.Cells(lastRow, indexExecutionColumnReducaoNCM)).Value
    
    ' Descobrir dimensões reais para a MatrizOut
    Dim rows As Long
    
    rows = UBound(RangeToMatriz, 1) ' Descobrindo o Limite susperior, ultimo index da linha na primeira coluna da matriz
    
    ' Redimensiona a Matriz criada com o Range com  saída de 5 colunas
    ReDim MatrizOut(1 To rows, 1 To 5)
    
    
    ' Agora vamos criar um iterador que pegue a quantidade de caracteres
    ' distribuia os caracteres pelas demais colunas a direita
    ' conforma a sua atribuicao: capitulo, posicao, subposicao "aqui pode ter genericos (5 caracteres)", itens, subitens
    For iRows = 1 To rows
    
        
     
    Next iRows
    
    
    ' Percorrendo cada celula
    For Each cel In rngNCM
        
        'valueCel = Trim(cel.Value)
        valueCel = cel.Value
        valueCel = Trim(valueCel)
        valueCel = Replace(valueCel, ".", "")
        valueCel = Replace(valueCel, ",", "")
        valueCel = Replace(valueCel, " ", "")
        
        ' -------------------------------------------------------
        ' 9 caracteres ? formato completo
        ' -------------------------------------------------------
        
        If valueCel <> "" And Len(valueCel) = 9 Then
            
            If regexNcmNove.Test(valueCel) Then
                Set match = regexNcmNove.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
            
            regexIdentificarServico = match.SubMatches(0)
            regexCapNcm = match.SubMatches(1)
            regexPosNcm = match.SubMatches(2)
            regexSubPosValue = match.SubMatches(3)
            regexItemValue = match.SubMatches(4)
            regexSubItemValue = match.SubMatches(5)
            
            cel.Offset(0, 1).Value = "Servico nao faz parte do Produto ou seja Sem NCM"
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""
            
            valueCel = regexIdentificarServico & "." & regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue & "." & regexSubItemValue
            
            cel.Value = CStr(valueCel)
            
            
        ' -------------------------------------------------------
        ' 8 caracteres ? formato completo
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 8 Then
        
            If regexNcmOito.Test(valueCel) Then
                Set match = regexNcmOito.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            regexItemValue = match.SubMatches(3)
            regexSubItemValue = match.SubMatches(4)
            
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue & "." & regexSubItemValue

            cel.Offset(0, 1).Value = "'" & CStr(regexCapNcm)
            cel.Offset(0, 2).Value = "'" & CStr(regexPosNcm)
            cel.Offset(0, 3).Value = "'" & CStr(regexSubPosValue)
            cel.Offset(0, 4).Value = "'" & CStr(regexItemValue)
            cel.Offset(0, 5).Value = "'" & CStr(regexSubItemValue)
            
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
            
        ' -------------------------------------------------------
        ' 7 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 7 Then
        
            If regexNcmSete.Test(valueCel) Then
                Set match = regexNcmSete.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            regexItemValue = match.SubMatches(3)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = "'" & regexSubPosValue
            cel.Offset(0, 4).Value = "'" & regexItemValue
            cel.Offset(0, 5).Value = ""
            
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
        ' -------------------------------------------------------
        ' 6 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 6 Then
        
            If regexNcmSeis.Test(valueCel) Then
                Set match = regexNcmSeis.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If

            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = "'" & regexSubPosValue
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""
                
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
        
        ' -------------------------------------------------------
        ' 5 caracteres -> Representando os gen?ricos
        ' -------------------------------------------------------
        
        ElseIf valueCel <> "" And Len(valueCel) = 5 Then
        
            If regexNcmCinco.Test(valueCel) Then
                    Set match = regexNcmCinco.Execute(valueCel)(0)
                Else
                    GoTo TrataElse
                End If

                regexCapNcm = match.SubMatches(0)
                regexPosNcm = match.SubMatches(1)
                regexSubPosValue = match.SubMatches(2)
            
                cel.Offset(0, 1).Value = "'" & regexCapNcm
                cel.Offset(0, 2).Value = "'" & regexPosNcm
                cel.Offset(0, 3).Value = "'" & regexSubPosValue
                cel.Offset(0, 4).Value = ""
                cel.Offset(0, 5).Value = ""
                
                valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue
            
                ' Formata tipo n?mero para string na exibi??o do Excel
                cel.NumberFormat = "@"
            
                cel.Value = CStr(valueCel)
                
        ' -------------------------------------------------------
        ' 4 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 4 Then
            
            ' Eu Guilherme Henrique n?o sei que Bruxaria ? Essa!
            'valueCel = Replace(valueCel, ",", ".")
            'valueCel = Replace(valueCel, ".", "")
            
            If regexNcmQuatro.Test(valueCel) Then
                Set match = regexNcmQuatro.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            
            
            valueCel = regexCapNcm & "." & regexPosNcm
            
           ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

        ' -------------------------------------------------------
        ' 2 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 2 Then
        
            If regexNcmDois.Test(valueCel) Then
                Set match = regexNcmDois.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

            cel.Value = CStr(valueCel)


        ' -------------------------------------------------------
        ' 1 caractere
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 1 Then
            
            valueCel = "0" & valueCel
            
            If regexNcmDois.Test(valueCel) Then
                Set match = regexNcmDois.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
        
            cel.Value = CStr(valueCel)
        
        Else
            GoTo TrataElse
        End If

        GoTo Proximo

TrataElse:
        cel.Offset(0, 1).Value = ""
        cel.Offset(0, 2).Value = ""
        cel.Offset(0, 3).Value = ""
        cel.Offset(0, 4).Value = ""
        cel.Offset(0, 5).Value = ""
    
        ' Formata tipo n?mero para string na exibi??o do Excel
        cel.NumberFormat = "@"
        
        cel.Value = CStr(valueCel)

Proximo:
    
    Next cel
    
    MsgBox "Planilha ReducaoNCM formatada com sucesso! Partes do codigo capturados com sucessos."
    
End Function

' ===========================================================
' FUNÇÕES AUXILIARES Para a Funcao CruzarNcm
' ===========================================================

Private Function SomenteDigitos(ByVal s As String) As String
    

End Function

Private Function GerarNiveisNCM(ByVal codigo As String) As Collection
'
End Function

Private Function ExisteChave(col As Collection, chave As String) As Boolean
'
End Function

Private Function BuildReductionCollection(ws As Worksheet) As Collection
'
End Function

Private Function CruzarNcmsPorNiveis(wsSrc As Worksheet, wsRed As Worksheet)
'
End Function


' Sub rotina publica auxilar flexivel que vamos usar para gravar os dados de qualquer matriz
Private Sub GravarInSheet( _
        ByVal nameSheet As String, _
        ByVal indexStart As Long, _
        ByVal indexColumn As Long)
    
    ' Criando uma Matriz de resultados com base na nossa funcao
    Dim resultado As Variant
    
    resultado = ReadFormatarNcmPlanOne(nameSheet, indexStart, indexColumn)
    
    ' materializnado um objeto planilha
    Dim Sheet As Worksheet
    
    ' usando a nossa planilha
    Set Sheet = Worksheets(nameSheet)
    
    ' Ativando a nossa planilha
    
    Sheet.Activate
    
    ' variavel que vai representar o numero da ultima linha
    Dim lastRow As Long
            
    lastRow = Sheet.Cells(Sheet.rows.Count, indexColumn).End(xlUp).Row

    Sheet.Range(Sheet.Cells(indexStart, indexColumn + 1), _
             Sheet.Cells(lastRow, indexColumn + 5)).Value = resultado
    
    MsgBox "Planilha NCM Formatada com Sucesso, bem como a distribuicao de cada item do codigo."

End Sub


Sub FormatarPlanilhasCruzarDados()
    
    
    Dim Sheet As Worksheet
    
    Dim RangeValidation As Range
    
    Dim UltimaCell As Long
    
    Application.ScreenUpdating = False ' desativando atualizacao visivel do codigo na planilha para ser mais rapido
            
    ' Aplicando Lógica de Verificação de formatação...
    
    ' Set Sheet = Worksheets(nameSheetPlanOne)
    
    Set Sheet = Worksheets("Itens das NF-es Recebidas - Aut")
    

        
    ' Atenção, se em algum momento a Planilha for modificada
    ' Deve-se modificar os index definidos aqui em cada objeto Cells
    
    UltimaCell = Sheet.Cells(Sheet.rows.Count, 12).End(xlUp).Row
    
    Set RangeValidation = Sheet.Range(Sheet.Cells(4, 8), Sheet.Cells(UltimaCell, 12))
    
     If ValidarRangeToFormat(RangeValidation) = False Then
        ' Range contém pelo menos uma célula vazia -> deve formatar
        Call GravarInSheet(nameSheetPlanOne, indexCellFirstColumn, indexColumnPlanOne)
    Else
        MsgBox "A planilha já esta na formatação adequada."
    End If

End Sub
