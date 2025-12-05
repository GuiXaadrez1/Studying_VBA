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

Private Function ReadRangeAsMatrix(intervalMatriz As String) As Range
    
    ' Lembrando que Matriz possui duas dimensoes, ou seja, são bidimensionais
    ' Ou seja, altura e largura, neste caso, uma tabela do Excel é uma matriz

     Set ReadRangeAsMatrix = Range(intevalMatriz)
        
     ReadRangeAsMatrix.Value = "Matriz Criada"
     
    Debug.Print "Foi Realizado a criação da Matriz com Range"
   
End Function


Sub ManipulationObjectRangeInCollectionsRangeWithMatriz()
    
     Call
    
End Sub

