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

' Option Base 1 -> opção ue faz todos os arrys começem com indicie 1 ao em vez de 0

' A opera??o direta c?lula-a-c?lula ? lenta porque cada acesso a uma c?lula envolve:

' Entrada/sa?da (I/O) com o Excel

' Sincroniza??o COM

' Atualiza??o visual / vol?til

' A solu??o profissional ?:

' Carregar o Range inteiro para uma matriz VBA, process?-la na mem?ria e devolv?-la ao Range.

Private Function ReadRangeAsMatrix(intervalMatriz As String) As Variant
    
    ' Lembrando que Matriz possui duas dimensoes, ou seja, s?o bidimensionais
    ' Ou seja, altura e largura, neste caso, uma tabela do Excel ? uma matriz

    ReadRangeAsMatrix = Range(intervalMatriz).Value
        
    ' ReadRangeAsMatrix.Value = "Matriz Criada"
     
    ' Debug.Print "Foi Realizado a cria??o da Matriz com Range"
   
End Function

Private Function LimitsInMatriz(ByVal Matriz As Variant)
    ' Definindo variáveis que serão a representação das nossas coordenadas, limites
    
    Dim LimitInferiorRow As Long ' Vai Retornar o primeiro indicie da primiera dimensão do array (Linha)
    Dim LimitSuperiorRow As Long  ' Vai Retornar o ultimo indicie da primiera dimensão do array (Linha)
    Dim LimitInferiorColumn As Long ' Vai Retornar o primeiro indicie da segunda dimensão do array (Coluna)
    Dim LimitSuperiorColumn As Long ' Vai Retornar o Ultimo indicie da segunda dimensão do array (Coluna)
    
    ' Definindo as variáveis que serão responsável pela nossa iteração na Matriz
    
    Dim i As Long
    Dim j As Long
    
    ' Lembrando que a interseção entre a linha e coluna é o que chamamos de célula aqui no excel
    ' Logo para atribuir uma célula com a Matriz formada
    ' Basta colocar = Matriz(i,j) = "Valor da Célula"
    
    ' Método LBound -> Retorna os limites inferiores de um array -> Primeiro indicie do array
    ' Método UBound -> Retorn os limites superiores de um array -> Ultimo indicie
    
    ' LBound(array, [dimension]) -> array que será examinado, dimesion -> opcional (1 para linhas, 2 para colunas)
    ' UBound(array, [dimension]) -> array que será examinado, dimesion -> opcional (1 para linhas, 2 para colunas)
    
    LimitInferiorRow = LBound(Matriz, 1)
    
    LimitSuperiorRow = UBound(Matriz, 1)
    
    LimitInferiorColumn = LBound(Matriz, 2)

    LimitSuperiorColumn = UBound(Matriz, 2)
    
    ' Irerando sobre o nosso array usando o Loop For
    
    For i = LimitInferiorRow To LimitSuperiorRow
        
        For j = LimitInferiorColumn To LimitSuperiorColumn
            
            ' Usando o objeto Cells paraprencher específicamente aquela célula da nossa matriz
            ' Durante essa iteração, usar a atribuição direta Matriz(i,j) = "Valor" é uma opção,
            ' mas com esse objeto é mais seguro
            
            Cells(i, j) = "Célula preenchida com Sucesso!"
            
            ' Em Debug.Print no VBA... o caractere ; serve para continuação
            ' Para termos quebra de linha no seu log do terminal ´´e necessário forçar usando a constante vbCrLf
            
            ' Debug.Print "Indicie inferior, Primeira Dimensão do Array: " & CStr(LimitInferiorRow) & vbCrLf & _
            '           "Indicie superior, Segunda Dimensao do Array: " & CStr(LimitSuperiorRow) & vbCrLf & _
            '            "Inidicie inferior, Segunda Dimensão do Array: " & CStr(LimitInferiorColumn) & vbCrLf & _
            '            "Indicie Superior, Segunda Dimensão do Array: " & CStr(LimitSuperiorColumn) & vbCrLf & _
            '            "Acessando a célula da linha: " & CStr(i) & " " & "Na coluna: " & CStr(j)
            
            Debug.Print "Acessando a célula da linha: " & CStr(i) & " " & "Na coluna: " & CStr(j)
            
        Next j
    
    Next i
    
End Function


Sub ManipulationObjectRangeInCollectionsRangeWithMatriz()
           
    Dim Matriz As Variant
    
    
    Matriz = ReadRangeAsMatrix("A1:B10")
    
    LimitsInMatriz (Matriz)

End Sub


