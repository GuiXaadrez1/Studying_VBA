Option Explicit

' Objetivo... Criar um C?digo que formate a Planilha e que use Collection para fazer um Cruazamento
' De informacoes por CNPJS/CPF, Se Na Planilha1 o Cnpj e igual ao Cnpj da Planilha2 Passar o elemento
' do Collection para celula ao lado do Cnpj da Planilha1


''''''''''''''''''''''''''''''''''''''''''
' Configuracao Gerais                    '
''''''''''''''''''''''''''''''''''''''''''

Const nameSheetThow As String = "Consta_SN"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Configuracao Para formatar Itens das NF-es Recebidas - Aut   '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const nameSheetOne As String = "Itens das NF-es Recebidas - Aut"

Const indexCabecalhoSomplesNacional As Integer = 3


Private Function FormatarPlanOne(ByVal nameSheet As String, Optional indexCellCacebalho As Integer = indexCabecalhoSomplesNacional) As Worksheet
    
    Dim i As Integer
    
    On Error GoTo Error
        
        Set FormatarPlanOne = Worksheets(nameSheet)
    
            ' Ativando a Planilha
            FormatarPlanOne.Activate
            
            ' Se a Primeira celula da primiera linha na coluna A for diferente de vazio, entao aplicar formatacao
            If FormatarPlanOne.Cells(1, 1) <> "" Then
            
                ' Inserindo duas linhas no topo
                For i = 1 To 2
                    FormatarPlanOne.rows("1:1").Insert
                Next i
                    
                ' Checa se os cabe?alhos j? est?o corretos
                If FormatarPlanOne.Cells(indexCellCacebalho, "C").Value = "Consta Simples Nacional" And _
                   FormatarPlanOne.Cells(indexCellCacebalho, "H").Value = "CAPITULO" Then
                   
                    ' Planilha j? formatada
                    MsgBox "A Planilha j? est? no formato correto!"
                    Exit Function
                End If
                
                ' --- Caso contr?rio, insere todas as colunas e cabe?alhos ---
                FormatarPlanOne.Columns("C").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "C").Value = "Consta Simples Nacional"
                
                FormatarPlanOne.Columns("H").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "H").Value = "CAPITULO"
                
                FormatarPlanOne.Columns("I").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "I").Value = "POSICAO"
                
                FormatarPlanOne.Columns("J").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "J").Value = "SUBPOSICAO"
                
                FormatarPlanOne.Columns("K").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "K").Value = "ITEM"
                
                FormatarPlanOne.Columns("L").Insert
                FormatarPlanOne.Cells(indexCellCacebalho, "L").Value = "SUBITEM"
                
                FormatarPlanOne.Columns("M").Insert
                
                ' Deixando da cor amarela
                FormatarPlanOne.Columns("M").Interior.Color = vbYellow
                
                FormatarPlanOne.Cells(indexCellCacebalho, "M").Value = "REDUCAO"
                
                MsgBox "A Planilha foi formatada com sucesso!"
                    
            Else
                MsgBox "A Planilha esta vazia ou ja esta no formato correto!"
                Exit Function
            End If
    
        ' Para evitar que cai no bloco Execucao de Error...
        Exit Function
            
Error:

    MsgBox "Nao foi Possivel Formatar a Primeira Planilha Devido Seguite Error: " & _
    vbCrLf & Err.Number & _
    vbCrLf & Err.Description
    
    Set FormatarPlanOne = Nothing ' Garante que a funcao retorne Nothing em caso de erro.

End Function

' Criando funcao que irá realizar a comparacao e cruzar CNPJS, Atencao, aqui coloquei o valor padrao com o nome da segunda planilha
' Para explicitar que e opcional
Private Function CruzarCnpjsWithCOllection(ByVal nameSheet As String, Optional sheetNameToCompare As String = nameSheetOne) As Variant
    
    Dim MyCollection As collection
    
    ' Materializando o Objeto da nossa Planilha
    Dim Sheet As Worksheet
        
    ' Materializando o Objeto da nossa Planilha que possui os cnpjs a serem comparados
    Dim SheetCompare As Worksheet
    
    ' Criando nosso Objeto que vai possuir nossa Matriz, Relativo a nossa Range
    Dim Matriz As Variant
    
    ' Defindo uma vari?vel que vai representar a nossa ultima linha preenchida encontrada naquela range
    Dim lastRow As Long
    
    Dim lastRowToSheetCompare As Long
    
    ' Materializando os nossos incrementadores para pecorrera nossa matriz criado atraves de um intervalo
    Dim i As Long
    ' Dim j As Long
    
    ' cnpj que vai ser comparado pela nossa Range Real e iterar sobre a mesma
    Dim cnpj As Variant
    
    ' Vamos usar a Range para fazer as compara??es e n?o a Matriz mais...
    'Dim rngCompare As Range
    'Dim cel As Range
           
    ' Vamos criar uma matriz atraves de uma range
    Dim RangeToMatriz As Variant
        
    ' Matriz de saida da nossa funcao que vai gravar os dados na planilha.
    Dim MatrizOut As Variant
    
    ' celula das coordenadas da nossa Matriz de saida
    Dim cel As Variant
        
    ' limites inferiores da nossa matriz de saida!
    Dim iRows As Long
    Dim iColumn As Long
     
    ' Materializando a nossa Collection
    Set MyCollection = New collection
    
    ' Referenciando e Materializando a Classe Worksheet
    Set Sheet = Worksheets(nameSheet)
     
    ' Ativando a nossa planilha
    Sheet.Activate
    
    lastRow = Sheet.Cells(Sheet.rows.Count, 1).End(xlUp).Row
    
    ' Adquirindo todos os objetos ao redor da nossa range que nao possui valor vazio
    ' Pegando a primeira celula e tambem a ultima celula preenchida para fazer o nosso Range... Intervalo de celula
    ' Matriz = Sheet.Range(Sheet.Cells(1, 1), Sheet.Cells(lastRow, 1)).CurrentRegion.Value -> funciona, mas deixei para fins didaticos
    
    Matriz = Sheet.Range(Sheet.Cells(1, 1), Sheet.Cells(lastRow, 2)).Value
    
    '  Desta dorma estamos pecorrendo cada linha e coluna, ou seja, a linha inteira e dentro dela cada Coluna
    ' LBound() + 1 -> Serve para pularmos o cabe?alho
    ' For i = LBound(Matriz, 1) + 1 To UBound(Matriz, 1)   ' percorre linhas
        
        ' Debug.Print "Acessando: "; CStr(i)
        ' Debug.Print "Acessando: "; CStr(Matriz(i, 1)) ' Acessando o valor da celula na primeira coluna
        
       ' MyCollection.Add(Key:=Matriz(i,1))
        
        ' For j = LBound(Matriz, 2) To UBound(Matriz, 2)   ' percorre colunas
            
            ' Debug.Print "Acessando: "; CStr(j)
            ' Debug.Print "Acessando: "; Matriz(i, j)
            ' Debug.Print "Acessando: "; CStr(Matriz(j, 2)) ' Acessando o valor da celula na segunda coluna
            
            ' MyCollection.Add(item:=Matriz(j,2))
       ' Next j
    ' Next i
    
    ' --------------------------------------------------------------------------------------------------------------
    
    ' Desta Forma estamos iterando sobre cada linha conforme a coordenada da linha em i e na coluna 1, coluna 2
    ' De uma s? vez, sem precisar pecorrer celula por celula individualmente, como com o duplo loop varia
    ' para todas as celulas da nossa matriz
    ' Com isso criamos um diconario simples com a Collection, ? o que queremos
    For i = LBound(Matriz, 1) + 1 To UBound(Matriz, 1)
    
        Dim chave As String
        Dim linha As Variant
        
        
        
        chave = NormalizeDoc(CStr(Matriz(i, 1)))   ' coluna A = chave
        linha = Matriz(i, 2)         ' coluna B = valor
        
        
        ' LEMBRANDO QUE COLLECTION NAO ACEITA CHAVES REPETIDAS! POR ISSO É NECESSARIO UMA VALIDACAO
        ' PARA ADICIONAR UMA CHAVE A ESSA COLLECTION
        
        ' se já existe, apaga
        If CollectionHasKey(MyCollection, chave) Then
            MyCollection.Remove chave
        End If
        
        ' adiciona o novo valor
        MyCollection.Add Item:=linha, key:=chave
        
        ' OBSERVACOES
        ' ------------------------------------------------------------
        ' O loop percorre *as linhas* da matriz retornada pelo Range.
        '
        ' Para cada linha, acessamos manualmente duas c?lulas:
        '   - Matriz(i, 1) -> valor da COLUNA 1 dessa linha (chave)
        '   - Matriz(i, 2) -> valor da COLUNA 2 dessa linha (valor)
        '
        ' Portanto, n?o existe loop sobre as colunas neste momento.
        ' Estamos apenas consultando duas colunas fixas (1 e 2) em cada iteracao.
        '
        ' A c?lula da COLUNA 1 representa a chave do item na Collection.
        ' A c?lula da COLUNA 2 representa o valor associado ? chave.
        '
        ' Isso transforma a Collection em um dicion?rio simples:
        '   chave -> valor
        '
        ' Ou seja, funcionalmente equivale a um array associativo:
        '   ("111" -> "Jo?o"), ("222" -> "Maria"), ...
        '
        ' Em resumo:
        '   COLUNA 1 -> chave da Collection
        '   COLUNA 2 -> valor da Collection
        ' ------------------------------------------------------------
        
        Debug.Print "Acessando o valor -> " & CStr(linha) & " -> Da Chave: " & CStr(chave)
    
    Next i
     
    ' Passando o Objeto Collectio por referencia para minha Funcao
    ' Set CruzarCnpjsWithCOllection = MyCollection
    
    Set SheetCompare = Worksheets(nameSheetOne)
    
    ' Ativadno a Sheet que possui os cnpjs a serem compados e cruzados.
    SheetCompare.Activate
    
    ' Pegando as ultima linha Preenchida... Igorando as duas ultimas que nao sao imprtantes para a comparacao
    lastRowToSheetCompare = SheetCompare.Cells(SheetCompare.rows.Count, "A").End(xlUp).Row - 2
       
    ' Criando uma Matriz de Saida com redimensionamento afim de fazer a comparacao
    
    'Set rngCompare = SheetCompare.Range(SheetCompare.Cells(4, 2), SheetCompare.Cells(lastRowToSheetCompare, 2))
    
    RangeToMatriz = SheetCompare.Range(SheetCompare.Cells(4, 2), SheetCompare.Cells(lastRowToSheetCompare, 3)).Value
    
    ' Iniciando a linha no primeiro index na primeira coluna
    For iRows = 1 To UBound(RangeToMatriz, 1)
        
        ' Acessando a ceula nas coordenadas... (index_linha,index_coluna)
        cel = RangeToMatriz(iRows, 1)
        
        ' o cnpj vai ser normalizado
        cnpj = NormalizeDoc(CStr(cel))
        
        ' Acessando o valor que esta na chave do meu cnpj
        If ExisteChave(MyCollection, CStr(cnpj)) Then
            
            cel = MyCollection.Item(cnpj)
            
             RangeToMatriz(iRows, 2) = cel
        Else
            RangeToMatriz(iRows, 2) = "Nao encontrado"
            
            'Debug.Print "Nao encontrado: "; cel
        End If
    
    Next iRows
          
    CruzarCnpjsWithCOllection = RangeToMatriz
          
End Function

' Criando funcao auxiliar para normalizar os cnpjs de comparacao
Private Function NormalizeDoc(ByVal valor As String) As String
    NormalizeDoc = Trim$(valor)
    NormalizeDoc = Replace(NormalizeDoc, vbTab, "")
    NormalizeDoc = Replace(NormalizeDoc, Chr(160), "") ' espaco nao-quebravel
End Function


' Criando funcao auxiliar para verificar se existe o valor na nossa collection
Private Function ExisteChave(col As collection, chave As String) As Boolean
    
    On Error GoTo ErrHandler
        Dim tmp As Variant
        
        tmp = col(chave)
        ExisteChave = True
        
        Exit Function

ErrHandler:
    ExisteChave = False
End Function


Private Function FormatCnpjAndCpfColumnASheetConstaSn( _
    Optional ByVal sheetNameToFormat As String = nameSheetThow)

    Dim Sheet As Worksheet
    Dim regexCPF As New RegExp
    Dim regexCNPJ As New RegExp
    Dim colValor As Range
    Dim linha As Long
    Dim valor As String
    Dim SomenteDigitos As String
    Dim match As Object
    Dim formatado As String

    ' --- Validacao da planilha ---
    On Error Resume Next
        Set Sheet = Worksheets(sheetNameToFormat)
    On Error GoTo 0

    If Sheet Is Nothing Then
        MsgBox "A planilha '" & sheetNameToFormat & "' nao existe.", vbCritical
        Exit Function
    End If

    ' --- Configuracao dos padroes ---
    regexCPF.Pattern = "^\d{11}$"
    regexCPF.Global = False

    regexCNPJ.Pattern = "^\d{14}$"
    regexCNPJ.Global = False

    linha = 2

    ' --- Loop ate encontrar c?lula vazia ---
    Do While Trim(Sheet.Cells(linha, 1).Value) <> ""

        valor = CStr(Sheet.Cells(linha, 1).Value)

        ' Remove qualquer caractere nao num?rico
        SomenteDigitos = Replace(valor, ".", "")
        SomenteDigitos = Replace(SomenteDigitos, "-", "")
        SomenteDigitos = Replace(SomenteDigitos, "/", "")
        SomenteDigitos = Replace(SomenteDigitos, " ", "")
        SomenteDigitos = Trim(SomenteDigitos)

        ' --- Normalizacao: completa com zeros ? esquerda ---
        If Len(SomenteDigitos) < 11 Then
            SomenteDigitos = String(11 - Len(SomenteDigitos), "0") & SomenteDigitos
        ElseIf Len(SomenteDigitos) > 11 And Len(SomenteDigitos) < 14 Then
            SomenteDigitos = String(14 - Len(SomenteDigitos), "0") & SomenteDigitos
        End If

        ' --- Formatacao CPF ---
        If regexCPF.Test(SomenteDigitos) Then
        
            formatado = _
                Mid$(SomenteDigitos, 1, 3) & "." & _
                Mid$(SomenteDigitos, 4, 3) & "." & _
                Mid$(SomenteDigitos, 7, 3) & "-" & _
                Mid$(SomenteDigitos, 10, 2)

        ' --- Formatacao CNPJ ---
        ElseIf regexCNPJ.Test(SomenteDigitos) Then
        
            formatado = _
                Mid$(SomenteDigitos, 1, 2) & "." & _
                Mid$(SomenteDigitos, 3, 3) & "." & _
                Mid$(SomenteDigitos, 6, 3) & "/" & _
                Mid$(SomenteDigitos, 9, 4) & "-" & _
                Mid$(SomenteDigitos, 13, 2)

        Else
            ' caso o valor nao seja cpf nem cnpj (mesmo ap?s limpeza)
            formatado = "Valor invalido"
        End If

        ' --- Escrita na celula ---
        Sheet.Cells(linha, 1).Value = formatado

        linha = linha + 1
    Loop

End Function

' FUNCAO AUXILIAR PARA COLLECTION DE CRUZAMENTO DE CNPJS

Public Function CollectionHasKey(col As collection, key As String) As Boolean
    Dim tmp As Variant
    
    On Error Resume Next
    tmp = col.Item(key)
    
    If Err.Number = 0 Then
        CollectionHasKey = True
    Else
        CollectionHasKey = False
        Err.Clear
    End If
    
    On Error GoTo 0
End Function


' CRIANDO UMA SUB AUXILIAR PARA A SUB PRINCIPAL, ELA VAI EXCUTAR A GRAVACAO DA MATRIZ PRONTA RETORNADA DA FUNCAO DE CRUZAMENTO
Public Sub GravarCruzamento( _
    ByVal nameSheet As String, _
    Optional ByVal sheetNameToCompare As String = nameSheetOne)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim resultado As Variant

    ' 1. Chama a função que monta a matriz de saída
    resultado = CruzarCnpjsWithCOllection(nameSheet, sheetNameToCompare)

    ' 2. Seleciona a planilha onde deve gravar os dados
    Set ws = Worksheets(sheetNameToCompare)

    ' 3. Descobre onde essa matriz deve ser escrita
    lastRow = ws.Cells(ws.rows.Count, "A").End(xlUp).Row - 2

    ' 4. Grava a matriz inteira de uma vez
    ws.Range( _
        ws.Cells(4, 2), _
        ws.Cells(lastRow, 3) _
    ).Value = resultado

    MsgBox "Cruzamento concluído com sucesso!", vbInformation

End Sub

Sub FormatarPlanilhaCruzarCnpjsAbaPrincipal()
    
    ' Definir a propriedade ScreenUpdating como False para ocultar as alteracoes enquanto elas sao feitas via codigo;
    ' Application.ScreenUpdating = True ' Deixei True porque com False eu nao sei se foir ou nao formatado...
    
    ' Chamando a nossa funcao para ser executada
    Call FormatarPlanOne(nameSheetOne)
    
    ' Formatando os Cnpjs e Cpfs
    MsgBox ("Clique em OK para fazer o Cruzamento dos CNPJS")
    
    Application.ScreenUpdating = False ' Agora Deixei False para agiilizar a execucao do c?digo
    
    ' Application.Interactive = False ' Impedindo do usuario mexer no excel para evitar erros inesperados
    
    Call FormatCnpjAndCpfColumnASheetConstaSn
    
    ' Application.ScreenUpdating = True ' Reativando para mostrar as modificacoes feitas
    ' Application.Interactive = True ' Reativando a interacao do usuario
    
    ' Chamando a nossa Funcao Para CruzarCNPJS com COllection
    ' Call CruzarCnpjsWithCOllection(nameSheetThow)


   Call GravarCruzamento(nameSheetThow, nameSheetOne)


End Sub