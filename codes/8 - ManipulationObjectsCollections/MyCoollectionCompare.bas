' ====================================================================================================
' CONFIGURAÇÃO PARTE CRUZAMENTO NCM
' ====================================================================================================

Const SourceSheetName As String = "Itens das NF-es Recebidas - Aut" ' Planilha com os NCMs de origem
Const SourceColumn As Long = 7 ' Coluna G: Coluna com os NCMs a serem lidos (Origem)
Const OutputColumn As Long = 13 ' Coluna M: Coluna onde o resultado da redução será inserido (Saída)
Const StartRowSource As Long = 4 ' Linha de início da iteração na planilha de origem

Const ReductionSheetName As String = "ReducaoNCM" ' Planilha com as taxas de redução
Const ReductionCodeNcmColumn As Long = 1 ' Coluna A: Coluna contendo os NCMs (Redução)
Const ReductionTaxColumn As Long = 7 ' Coluna G: Coluna contendo a taxa de redução (Redução)
Const StartRowReduction As Long = 2 ' Linha de início da iteração na planilha de redução


' ===========================================================
' FUNÇÕES AUXILIARES
' ===========================================================

' Padroniza NCM, removendo pontos e espaços para garantir compatibilidade com a collection de redução.
Private Function NormalizarNCM(ByVal codigo As String) As String
    Dim ncmClean As String
    ncmClean = Replace(Replace(Trim(codigo), ".", ""), " ", "")
    
    ' Usa RegEx para remover todos os caracteres que não são dígitos (mais eficiente)
    Static regex As Object

    If regex Is Nothing Then
        
        ' Materializando um Objeto regex
        Set regex = CreateObject("VBScript.RegExp")
        
        With regex
            .Pattern = "\D" ' Qualquer não-dígito -> Exmeplo: Abc34as124, vai retornar: 34124
            .Global = True
        End With
    
    End If
    
    NormalizarNCM = regex.Replace(ncmClean, "")

End Function

' Gera os níveis de NCM para busca: 8, 7, 6, 5, 4, 2 dígitos.
Private Function GerarNiveisNCM(ByVal codigoNCM As String) As collection
    
    ' Observação importante:
    
    ' Essa funcao possui uma collection com o prorposito de armazena os niveis do NCM
    ' ou seja, os ncms vao ser itens dessa collection
    
    ' Exemplo: Para o NCM 12345678, a collection terá:
    
    ' item: 12345678 de key 1 (índice 1 da collection) 
    ' item: 1234567 de key 2 (índice 2 da collection)
    ' item: 123456 de key 3 (índice 3 da collection)
    ' item: 12345 de key 4 (índice 4 da collection)
    ' item: 1234 de key 5 (índice 5 da collection)
    ' item: 12 de key 6 (índice 6 da collection)

    ' Observação Dois!

    ' Eu nao sabia, mas collection permite valores duplicados! 
    ' Mas ele nao permite chaves duplicadas. (Key -> Unico, Item -> pode repetir)
    ' Porque a Key é o identificador unico do item na collection.

    Dim ncmNormalizado As String
   
    Dim collectionData As New collection

    ' Atribuindo implicitamente a nova collection ao retorno da função 
    Set GerarNiveisNCM = collectionData
    
    ncmNormalizado = NormalizarNCM(codigoNCM)
    
    ' Adiciona os níveis do mais específico para o mais genérico
    If Len(ncmNormalizado) >= 8 Then collectionData.Add Left$(clean, 8)
    If Len(ncmNormalizado) >= 7 Then collectionData.Add Left$(clean, 7)
    If Len(ncmNormalizado) >= 6 Then collectionData.Add Left$(clean, 6)
    If Len(ncmNormalizado) >= 5 Then collectionData.Add Left$(clean, 5) ' representa os códigos genericos
    If Len(ncmNormalizado) >= 4 Then collectionData.Add Left$(clean, 4)
    If Len(ncmNormalizado) >= 2 Then collectionData.Add Left$(clean, 2)
    
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
    Optional ColumnCodeNcmVariat As Long = ReductionCodeNcmColumn, _
    Optional ColumnTaxRedution As Long = ReductionTaxColumn _
) As collection
    
    ' Observação importante:
    ' Essa funcao possui uma collection com o proposito de armazenar as taxas de reducao
    ' onde o codigo NCM sera a chave (key) e o valor da reducao sera o item.
    ' Logo os ncms sao unicos nessa collection e nao podem se repetir.
    ' vamos usar os ncms como chave para facilitar a busca e as comparacoes.
    ' Exemplo: Para o NCM 12345678, a collection terá:
    ' key: 12345678 -> item: 0.10 (10% de redução)

    ' ultima linha preenchida na planilha de redução
    Dim lastRow As Long

    ' Matriz para carregar os dados da planilha
    Dim RangeToMatriz As Variant

   ' Contador de linhas na matriz 
    Dim iRows As Long
    
    ' Vai armazenar o código NCM e o valor de redução
    Dim codigoNcm As String
    Dim valorReducao As Variant

    ' Se o nome da planilha não for passado, usa a constante
    If Sheet Is Nothing Then Set Sheet = Worksheets(ReductionSheetName)

    ' Atribuindo implicitamente a nova collection ao retorno da função
    Set BuildReductionCollection = New collection

    ' Última linha da coluna de códigos NCM    
    lastRow = Sheet.Cells(Sheet.rows.Count, ColumnCodeNcmVariat).End(xlUp).Row
    
    ' Carrega a área de dados em uma matriz para velocidade
    RangeToMatriz = Sheet.Range( _
                        Sheet.Cells(StartCellIndex, ColumnCodeNcmVariat), _
                        Sheet.Cells(lastRow, ColumnTaxRedution) _
                      ).Value
    
    
    ' A matriz RangeToMatriz é 1-based e tem UBound(matriz, 2) colunas
    Dim codeNcmColIndex As Long
    Dim taxReductionColIndex As Long
    
    ' Calcula os índices das colunas dentro da matriz (diferença entre colunas do Excel)
    codeNcmColIndex = 1 ' Sempre será a primeira coluna da matriz carregada
    
    ' Calcula o índice da coluna de redução dentro da matriz carregada
    taxReductionColIndex = ColumnTaxRedution - ColumnCodeNcmVariat + 1
    
    ' resultado do calculo: se ColumnTaxRedution = 7 e ColumnCodeNcmVariat = 1
    ' taxReductionColIndex = 7 - 1 + 1 = 7 (sétima coluna da matriz)

    ' UBound -> Retorna o maior índice de uma matriz em uma dimensão especificada
    For iRows = 1 To UBound(RangeToMatriz, 1)
        
        ' Normaliza o código NCM como chave para garantir compatibilidade e facilitar a comparação
        codigoNcm = NormalizarNCM(CStr(RangeToMatriz(iRows, codeNcmColIndex)))
        
        ' Obtém o valor de redução correspondente
        valorReducao = RangeToMatriz(iRows, taxReductionColIndex)
        
        If Len(codigoNcm) > 0 Then
            'On Error evita que códigos duplicados causem erro.
            ' Basicamente vamos ignorar a duplicata
            On Error Resume Next
                ' Adiciona a redução o valor da  reducao na collection.
                ' Conforme a sua chave (codigNcm -> normalizado)
                ' Para realizarmos comparacao 
                BuildReductionCollection.Add Item:=valorReducao, key:=codigoNcm
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
    
    ' Matrizes para entrada, a que vamos comparar
    Dim matrizFonte As Variant

    ' Matriz de saída (resultado)
    Dim matrizSaida() As Variant
    
    ' Collection retornado pela funcao BuildReductionCollection
    Dim reductionCollection As collection
    
    ' Collection para armazenar os níveis do NCM, gerenciar os mesmos
    Dim niveisCollection As collection
      
    Dim codNcm As String
    
    ' Representa cada nível do NCM (8,7,6,5,4,2) que usamos para armazenar
    ' no collection GerarNiveisNCM e usar cada nivel desta collection para comparar
    ' com a collection reductionCollection

    ' Ou seja, primeiro é o ncm como item 
    ' O segundo e passando o ncm como item para comparar com a chave da collection reductionCollection
    ' Se for correspondente, pegamos o valor da reducao (item) e armazenamos na matriz de saida.

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
    
    ' UBound(matrizFonte, 1) → retorna o número máximo de linhas da matrizFonte
    ' 1 To 1 → significa que a matriz de saída terá uma coluna
    ' Isso é o mesmo que: matrizSaida(UBound(matrizFonte,1),1)
    ' Cada linha da matrizSaida corresponde exatamente a uma linha da matrizFonte
    ' Como a Matriz Fonte Ja existe, Criamos a Matriz Saida com o mesmo numero de linhas
    ' Porem, ela esta vazia, aguardando os resultados do cruzamento.

    ' Collection da planilha ReducaoNCM (carregada uma vez)
    Set reductionCollection = BuildReductionCollection(Worksheets(ReductionSheetName))
    
    ' Loop principal para processar cada NCM na matriz de entrada
    For iRows = 1 To UBound(matrizFonte, 1)
        
        ' Codigo NCM normalizado -> valor da celula atual da matrizFonte
        ' normalizado e tansformada em string

        codNcm = NormalizarNCM(CStr(matrizFonte(iRows, 1)))

        ' Se o NCM estiver vazio, resultado é 0%        
        If Len(Trim(codNcm)) = 0 Then
            matrizSaida(iRows, 1) = "0%"
            ' Vai para o próxima iteração do loop
            GoTo Proximo
        End If
        

        ' Collection que Gerencia os níveis para o NCM (8, 7, 6, 5, 4, 2)
        ' lembrando que ele vai fazer a validacao do codigo ncm do 
        ' mais especifico para o mais generico
        ' exemplo: 12345678 -> 12345678, 1234567, 123456, 12345, 1234, 12
        ' perceba que o codigo esta sem pontos e espacos, para facilitar a comparacao


        ' Apos ter passado pela validavao, passamos o NCM da MatrizFonte
        ' para a funcao GerarNiveisNCM, que vai retornar uma collection
        ' com os niveis do NCM (8,7,6,5,4,2)
        Set niveisCollection = GerarNiveisNCM(codNcm)
        
        ' de inicio, nao encontrou a reducaom, logo é false
        found = False

        ' valor padrao caso nao encontre a reducao é 0%
        resultado = "0%"
        
        ' Testa nivel por nível (do mais específico ao mais genérico)
        ' o nivel é cada item da collection niveisCollection (gerada pela funcao GerarNiveisNCM) 
        For Each nivel In niveisCollection
            
            ' passando o nivel (item da collection niveisCollection) como string 
            ' para comparar com a chave da collection reductionCollection

            ' ExisteChave -> Funcao que valida se a chave existe na collection 
            If ExisteChave(reductionCollection, CStr(nivel)) Then

                ' Se existe o nivel na collection de reducao, pega o valor da reducao  
                resultado = reductionCollection(CStr(nivel))
                
                ' retorna verdadeiro, pois encontrou a reducao
                found = True
                
                ' Sai do loop, pois ja encontrou a reducao
                Exit For 
            End If
        
            ' Se nao encontrou, continua para o proximo nivel (item da collection niveisCollection)
        Next nivel
        
        ' Se encontrou a redução, formata o resultado como porcentagem
        If found Then
            ' Formata para porcentagem (resultado é um decimal, p.ex., 0.10)
            matrizSaida(iRows, 1) = Format(CDbl(resultado), "0%")
        Else
            ' Se não encontrou, mantém como "0%"
            matrizSaida(iRows, 1) = "0%"
        End If
                
Proximo:
        ' Próxima linha na matriz de entrada
    Next iRows
    
    ' Apos processar todos os NCMs, retorna a matriz de saída completa
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
