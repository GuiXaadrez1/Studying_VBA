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

' A operação direta célula-a-célula é lenta porque cada acesso a uma célula envolve:

' Entrada/saída (I/O) com o Excel

' Sincronização COM

' Atualização visual / volátil

' A solução profissional é:

' Carregar o Range inteiro para uma matriz VBA, processá-la na memória e devolvê-la ao Range.

Private Function ReadRangeAsMatrix(intervalMatriz As String) As Variant
    
    ' Lembrando que Matriz possui duas dimensoes, ou seja, são bidimensionais
    ' Ou seja, altura e largura, neste caso, uma tabela do Excel é uma matriz

    ReadRangeAsMatrix = Range(intervalMatriz).Value
        
    ' ReadRangeAsMatrix.Value = "Matriz Criada"
     
    ' Debug.Print "Foi Realizado a criação da Matriz com Range"
   
End Function


Sub ManipulationObjectRangeInCollectionsRangeWithMatriz()
    
    Dim Matriz As Variant
    
    ' Descobrindo os limites (dimensoões) da Matriz
    
    Dim rOne As Long
    Dim rTwo As Long
    Dim cOne As Long
    Dim cTwo As Long
    
    ' Definido iteradores para pecorrer a matriz
    
    Dim i As Long
    Dim j As Long
    
    ' Craindo a nossa Matriz através do Range
    Matriz = ReadRangeAsMatrix("A1:C1000")
    
    ' O VBA cria um array bidimensional com limites:
    ' Forma de criar Matriz Com VBA:
    ' Matriz(1 To 10, 1 To 3) -> primeiro representa as linhas de 1 até 10 e o segunda as colunas de 1 ate 3
    ' Lembrando que esses numeros são indexs...
    
    ' Isso permite trabalhar com QUALQUER tamanho de Range.
    rOne = LBound(Matriz, 1)
    rTwo = UBound(Matriz, 1)
    cOne = LBound(Matriz, 2)
    cTwo = UBound(Matriz, 2)
    
    ' Iterando sobre a nossa Matriz e Aplicando valores em suas células
    ' Perceba...
    
    For i = rOne To rTwo
        For j = cOne To cTwo
            
            ' Matriz(i, j) = "Matriz Preenchida com  Sucesso!"
            
            ' Cells é um objeto collection Cells
            
            Cells(i, j).Value = "Matriz Preenchida com  Sucesso!"
            
            Debug.Print "Linha " & i & " Coluna " & j & " = " & Matriz(i, j)
        Next j
    Next i
    
End Sub

