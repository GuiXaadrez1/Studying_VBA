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

' A opera��o direta c�lula-a-c�lula � lenta porque cada acesso a uma c�lula envolve:

' Entrada/sa�da (I/O) com o Excel

' Sincroniza��o COM

' Atualiza��o visual / vol�til

' A solu��o profissional �:

' Carregar o Range inteiro para uma matriz VBA, process�-la na mem�ria e devolv�-la ao Range.

Private Function ReadRangeAsMatrix(intervalMatriz As String) As Variant
    
    ' Lembrando que Matriz possui duas dimensoes, ou seja, s�o bidimensionais
    ' Ou seja, altura e largura, neste caso, uma tabela do Excel � uma matriz

    ReadRangeAsMatrix = Range(intervalMatriz).Value
        
    ' ReadRangeAsMatrix.Value = "Matriz Criada"
     
    ' Debug.Print "Foi Realizado a cria��o da Matriz com Range"
   
End Function


Sub ManipulationObjectRangeInCollectionsRangeWithMatriz()
    
    Dim Matriz As Variant
    
    ' Descobrindo os limites (dimenso�es) da Matriz
    
    Dim rOne As Long
    Dim rTwo As Long
    Dim cOne As Long
    Dim cTwo As Long
    
    ' Definido iteradores para pecorrer a matriz
    
    Dim i As Long
    Dim j As Long
    
    ' Craindo a nossa Matriz atrav�s do Range
    Matriz = ReadRangeAsMatrix("A1:C1000")
    
    ' O VBA cria um array bidimensional com limites:
    ' Forma de criar Matriz Com VBA:
    ' Matriz(1 To 10, 1 To 3) -> primeiro representa as linhas de 1 at� 10 e o segunda as colunas de 1 ate 3
    ' Lembrando que esses numeros s�o indexs...
    
    ' Isso permite trabalhar com QUALQUER tamanho de Range.
    rOne = LBound(Matriz, 1) ' Pega o limite inferior da primeira dimens�o (linhas)
    rTwo = UBound(Matriz, 1) ' Pega o limite superior da primeira dimens�o (linhas)
    cOne = LBound(Matriz, 2) ' Pega o limite inferior da segunda dimens�o (colunas)
    cTwo = UBound(Matriz, 2) ' pega o limite superior da segunda dimens�o (colunas)
    
    ' Iterando sobre a nossa Matriz e Aplicando valores em suas c�lulas
    ' Perceba...
    
    For i = rOne To rTwo
        For j = cOne To cTwo
            
            ' Matriz(i, j) = "Matriz Preenchida com  Sucesso!"
            
            ' Cells � um objeto collection Cells
            
            Cells(i, j).Value = "Matriz Preenchida com  Sucesso!"
            
            Debug.Print "Linha " & i & " Coluna " & j & " = " & Matriz(i, j)
        Next j
    Next i 
    
End Sub

