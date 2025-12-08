VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Objetivo... Criar um C�digo que formate a Planilha e que use Collection para fazer um Cruazamento
' De informa��es por CNPJS/CPF, Se Na Planilha1 o Cnpj � igual ao Cnpj da Planilha2 Passar o elemento
' do Collection para c�lula ao lado do Cnpj da Planilha1


''''''''''''''''''''''''''''''''''''''''''
' Configura��o Gerais                    '
''''''''''''''''''''''''''''''''''''''''''

Const nameSheetThow As String = "Consta_SN"

''''''''''''''''''''''''''''''''''''''''''
' Configura��o Para formatar Planilha1   '
''''''''''''''''''''''''''''''''''''''''''

Const nameSheetOne As String = "Planilha1"

Const indexCabecalhoSomplesNacional As Integer = 3


Private Function FormatarPlanOne(ByVal nameSheet As String, Optional indexCellCacebalho As Integer = indexCabecalhoSomplesNacional) As Worksheet
    
    Dim i As Integer
    
    On Error GoTo Error
        
        Set FormatarPlanOne = Worksheets(nameSheet)
    
            ' Ativando a Planilha
            FormatarPlanOne.Activate
            
            ' Se a Primeira celula da primiera linha na coluna A for diferente de vazio, entao aplicar formatacao
            If FormatarPlanOne.Cells(1, 1) <> "" Then

                ' Inserindo tres linhas
                
                For i = 1 To 2
                    ' A cada itera��o, insere uma nova linha na posi��o 1, empurrando as demais
                    FormatarPlanOne.Rows("1:1").Insert
                Next
            
        
                If FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "C").Value <> "" And FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "C").Value = "Consta Simples Nacional" Then
                 
                    ' Perfeito... Isso significa que existe um valor na celula entao
                    ' nao vai ser necessario colocar um valor nesta celula em especifico para cabecalho
                    
                    Exit Function
                
                ElseIf FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "C").Value = "" Then
                    
                    ' Inserindo a coluna C
                    FormatarPlanOne.Columns("C").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna C
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "C").Value = "Consta Simples Nacional"
                    
                    ' Inserindo a coluna H
                    FormatarPlanOne.Columns("H").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna H
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "H").Value = "CAPITULO"
                    
                    ' Inserindo a coluna I
                    FormatarPlanOne.Columns("I").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna I
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "I").Value = "POSICAO"
                    
                    ' Inserindo a coluna J
                    FormatarPlanOne.Columns("J").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna J
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "J").Value = "SUBPOSICAO"
                
                    ' Inserindo a coluna K
                    FormatarPlanOne.Columns("K").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna K
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "K").Value = "ITEM"
                    
                    ' Inserindo a coluna L
                    FormatarPlanOne.Columns("L").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna L
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "L").Value = "SUBITEM"
                    
                    ' Inserindo a coluna M
                    FormatarPlanOne.Columns("M").Insert
                
                    ' Inserindo Cabelha�o na terceira linha na coluna M
                    FormatarPlanOne.Cells(indexCabecalhoSomplesNacional, "M").Value = "REDUCAO"
                
                    MsgBox "A Planilha formatada com Sucesso!"
                
                End If
            
            Else
                
                MsgBox "A Planilha ja foi formatada, ou j� esta no formato Correto!"
                ' Usando uma Boa Pr�tica Para obrigar o c�digo a sair da fun��o assim que ele Cria a Planilha
                
                Exit Function
            End If
            
        ' Para evitar que cai no bloco Execu��o de Error...
        Exit Function
Error:

    MsgBox "N�o foi Possivel Formatar a Primeira Planilha Devido Seguite Error: " & _
    vbCrLf & Err.Number & _
    vbCrLf & Err.Description
    
    Set FormatarPlanOne = Nothing ' Garante que a fun��o retorne Nothing em caso de erro.

End Function

Private Function CruzarCnpjsWithCOllection(ByVal nameSheet As String)
    
    Dim MyCollection As Collection
    
    ' Materializando o Objeto da nossa Planilha
    Dim Sheet As Worksheet
    
    ' Criando nosso Objeto que vai possuir nossa Matriz, Relativo a nossa Range
    Dim Matriz As Variant
    
    ' Defindo uma vari�vel que vai representar a nossa ultima linha preenchida encontrada naquela range
    Dim lastRow As Long
    
    ' Materializando os nossos incrementadores para pecorrera nossa matriz criado atraves de um intervalo
    Dim i As Long
    Dim j As Long
    
    ' Materializando a nossa Collection
    Set MyCollection = New Collection
    
    ' Referenciando e Materializando a Classe Worksheet
    Set Sheet = Worksheets(nameSheet)
     
    ' Ativando a nossa planilha
    Sheet.Activate
    
    lastRow = Sheet.Cells(Sheet.Rows.Count, 1).End(xlUp).Row
    
    ' Adquirindo todos os objetos ao redor da nossa range que n�o possui valor vazio
    ' Pegando a primeira celula e tamb�m a ultima celula preenchida para fazer o nosso Range... Intervalo de celula
    Matriz = Sheet.Range(Sheet.Cells(1, 1), Sheet.Cells(lastRow, 1)).CurrentRegion.Value
    
    '  Desta dorma estamos pecorrendo cada linha e coluna, ou seja, a linha inteira e dentro dela cada Coluna
    ' LBound() + 1 -> Serve para pularmos o cabe�alho
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
    ' De uma s� vez, sem precisar pecorrer celula por celula individualmente, como com o duplo loop varia
    ' para todas as celulas da nossa matriz
    ' Com isso criamos um diconario simples com a Collection, � o que queremos
    For i = LBound(Matriz, 1) + 1 To UBound(Matriz, 1)
    
        Dim chave As String
        Dim linha As Variant

        chave = CStr(Matriz(i, 1))   ' coluna A = chave
        linha = Matriz(i, 2)         ' coluna B = valor
    
        MyCollection.Add Item:=linha, key:=chave
        
        ' OBSERVA��ES
        ' ------------------------------------------------------------
        ' O loop percorre *as linhas* da matriz retornada pelo Range.
        '
        ' Para cada linha, acessamos manualmente duas c�lulas:
        '   - Matriz(i, 1) -> valor da COLUNA 1 dessa linha (chave)
        '   - Matriz(i, 2) -> valor da COLUNA 2 dessa linha (valor)
        '
        ' Portanto, n�o existe loop sobre as colunas neste momento.
        ' Estamos apenas consultando duas colunas fixas (1 e 2) em cada itera��o.
        '
        ' A c�lula da COLUNA 1 representa a chave do item na Collection.
        ' A c�lula da COLUNA 2 representa o valor associado � chave.
        '
        ' Isso transforma a Collection em um dicion�rio simples:
        '   chave -> valor
        '
        ' Ou seja, funcionalmente equivale a um array associativo:
        '   ("111" -> "Jo�o"), ("222" -> "Maria"), ...
        '
        ' Em resumo:
        '   COLUNA 1 -> chave da Collection
        '   COLUNA 2 -> valor da Collection
        ' ------------------------------------------------------------
        
        Debug.Print "Acessando o valor -> " & CStr(linha) & " -> Da Chave: " & CStr(chave)
    
    Next i
     
    ' Passando o Objeto Collectio por refer�ncia para minha Funcao
    Set CruzarCnpjsWithCOllection = MyCollection
    
End Function

Sub ManipulationCollectionVBA()
    
    ' Chamando a nossa funcao para ser executada
    Call FormatarPlanOne(nameSheetOne)
    
    ' Chamando a nossa Funcao Para CruzarCNPJS com COllection
    Call CruzarCnpjsWithCOllection(nameSheetThow)

End Sub


' COLOCANDO ESSA SUB COMO REFERENCIA!

Sub CruzarCNPJs()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim linha As Long
    Dim cnpj As String
    
    Set ws = Worksheets("PlanilhaQueVouCruzar")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Percorre cada linha da planilha destino
    For linha = 2 To ultimaLinha
        
        cnpj = CStr(ws.Cells(linha, "A").Value) ' CNPJ da linha
        
        ' Verifica se a chave existe na Collection
        On Error Resume Next
        Dim valor As Variant
        valor = MyCollection(cnpj) ' tenta pegar o valor pela chave
        On Error GoTo 0
        
        ' Se encontrou a chave
        If Not IsEmpty(valor) Then
            ' Aqui vamos Usar o Método Offset(), vai ser mais fácil de entender e aplicar
            ws.Cells(linha, "B").Value = valor
        Else
            ws.Cells(linha, "B").Value = ""   ' ou "N�o encontrado"
        End If
        
        valor = Empty
        
    Next linha

End Sub

