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

' O objeto Cells representa uma referência matricial dentro do Excel. Em termos formais:

' Ele é um indexador: Cells(linha, coluna)

' Retorna um Range que representa exatamente uma célula.

' Quando usado sem argumentos, ele retorna todas as células da planilha ativa, funcionando como uma matriz virtual
' de dimensão igual à última célula utilizada do Excel (1.048.576 × 16.384).

' Composição interna

' Cells é matematicamente um tensor bidimensional (linha × coluna).

' Ou seja: Cells(i, j)  =  Range(row_index = i, column_index = j)

' E cada instância é um RANGE unitário (1 célula só), com:

' .Value
    
' .Formula

' .Interior

' .Font

' .Address

' Propriedades herdadas do Range.


Private Function CellsSemQualificador(ByVal index_row As Long, index_column As Long) As Object
    
    ' Se == Cells(1,1) -> Ele vai pegar exatamente a célula da primeira linha na primeira coluna dentro
    ' de um range... no caso a planilha toda!
    ' Mas isso só vai funcionar para a planilha ativa, se você estiver trabalhando em outra planilha
    ' Pode modificar o resultado! O ideal é sempre indicar a planilha correta para não dar problema...
    
   On Error Resume Next
    
    Set CellsSemQualificador = Cells(index_row, index_column)
   
    CellsSemQualificador.Value = "Celula Definida Preenchida"
   
End Function

Private Function CellsComQualificador(ByVal nameSheet As String, index_row As Long, index_column As Long) As Object
    
    Set CellsComQualificador = Worksheets(nameSheet).Cells(index_row, index_column)
    
    CellsComQualificador.Value = "Celula2 Preenchida!"

End Function


Private Function CellsInRange(ByVal nameSheet As String, interval As String, index_row As Long, index_column As Long) As Object
    
    Set CellsInRange = Worksheets(nameSheet).Range(interval).Cells(index_row, index_column)
    
    CellsInRange.Value = "Celula Preenchida Relativamente ao Intervalo (Range) de Celulas"

End Function


Private Function FindLastRowWithCells()
    
    Dim lastRow As Long
    
    ' Vamos pega a ultima célula, deslocar para a proxima célula e preencher
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
    ' Rows -> Objeto Row vamos ver mais no futuro como funciona
    ' Rows.Count -> Retorna o total de linhas da nossa planilha ativa
    ' .End(xlUp) -> erve para simular o ato de pressionar a tecla <kbd>End</kbd> e depois uma das teclas direcionais (<kbd>?</kbd>, <kbd>?</kbd>, <kbd>?</kbd>, <kbd>?</kbd>) no Excel.
    ' neste caso estamos pegando a ultima célula preenchida do excel com xlUp, famoso setinha para cima
    
    ' Usar Cells para referenciar a célula na nova linha encontrada (Coluna 1, Linha lastRow)
    Cells(lastRow, 1).Value = "Celula Preenchida Após a ultima celula Preenchida ser encontrada."
    
End Function

Private Function FindLastColumnCells()
    
    ' Vai seguir a Mesma lógica da funcao -> FindLastRowWithCells
    Dim lastColumn As Long
    
    ' Neste caso vamos encontrar ultima coluna preenchida, dislocar para a proxima coluna e preencher a celula
    ' na posicao especifica usando o objeto Cells
    lastColumn = Cells(1, Columns.Count).End(xlToLeft).Column + 1
    
    Cells(1, lastColumn).Value = "Celula Preenchida apos a ultima coluna preenchida encontrada."

End Function

Sub ManipulationObjectCellsWithVBA()
    
    ' Representa a célula retornada na nossa primiera função
    Dim CellOne As Object
    
    ' Representa a célula retornada na segunda função
    Dim Celltwo As Object
    
    Dim CellTree As Object
    
    On Error Resume Next
        ' Lembrando que esse formato so funciona para a planilha ativa
        ' CellOne = CellsSemQualificador(2, 5)
        
        ' Agora estamos indicando explícitamente qual planilha estamos trabalhando e em qual célula dentro dela vamos preencher
        ' Essa é a forma mais indicada para fazer manipulação diretamente na célula. (BOA PRÁTICA)
        Celltwo = CellsComQualificador("Planilha2", 2, 2)
          
        CellTree = CellsInRange("Planilha1", "A1:A10", 5, 1) ' -> vai acessar a célula relativo ao nosso intervalo de celula
        
        ' CellTree = CellsInRange("Planilha1", "A1:A10", 1, 5) -> veja que ele vai colocar a celula na quinta coluna, mas realtivo ao nosso range
                
        CellTree = CellsInRange("Planilha1", "A1:A10", 11, 1) ' Mesmo conceito do de cima, ao invés de colunas é uma célula a mais...
                
    On Error GoTo 0
    
        FindLastRowWithCells
        
        FindLastColumnCells
        
End Sub


