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

' O objeto Columns � um conjunto index�vel de objetos Range,
' onde cada elemento corresponde a uma coluna inteira da planilha ou do Range referenciado.

' Formalmente:

' Columns(n)  ->  Range que representa a coluna n inteira
' Columns("C") -> Range da terceira coluna

' E:

' Worksheet.Columns -> cole��o completa de colunas da planilha (Colunas Relativas ao Planilha Ativa)
' Range.Columns -> cole��o das colunas internas ao Range (Colunas Relativas ao Intervalo de Celulas )

' A DIFEREN�A FUNDAMENTAL ENTRE Columns E Cells

' Cells -> Objeto matricial bidimensional. Sempre retorna uma �nica c�lula.

' Columns -> Objeto unidimensional de faixa vertical. Sempre retorna um Range vertical, da linha 1 � 1.048.576.

Private Function ColumnRelativeSheet(ByVal nameSheet As String, Optional index_column As Variant = 1)
        
        ' Lembrando que o index da coluna pode ser em String ou seja, o nome da coluna
        Set ColumnRelativeSheet = Worksheets(nameSheet).Columns(index_column)
        
        ColumnRelativeSheet.Interior.Color = vbYellow

End Function


' Usamos o Optional da fun��o para indicar que � um parametro n�o obrigatorio,
' Geralmente nao precisamos colocar valor padrao no tipo Variant, mas nas demais, precisa-se
' Ou declarar em algum canto internamente na funcao

Private Function ColumnRelativeObjectRange(ByVal interval As String, Optional index_column As Variant = 1)
    
    ' Pegando a coluna relatica a nossa Range
    ' Lembrando novamente que o index da column pode ser tambem o nome da propria coluna, exemplo D
    
    Set ColumnRelativeObjectRange = Range(interval).Columns(index_column)
    
    MsgBox "Enrede�o da Coluna: " & CStr(ColumnRelativeObjectRange.Address)

End Function


Private Function ColumnCounts(ByVal interval As String) As Range
    
    ' o metodo count sem coontexto da Planilha ou Range, vai ser contabilizado todas as colunas da planilha ativa
    ' Com isso ele retorna o total de colunas existentes naquela planilha
    
    ' Contabilizando a quantidade de colunas existentes relativamente a um intervalo de celulas
    Set ColumnCounts = Range(interval)
    
    MsgBox "Quantidade Total de colunas neste intervalo de celulas: " & ColumnCounts.Columns.Count
    
    Exit Function
    
End Function

Sub ManipulationObjectColumnWithVBA()
    
    Dim ColumnColoridaOne As Object
    
    Dim AddressColumn As Object
    
    Set ColumnColoridaOne = ColumnRelativeSheet("Planilha1", 2)
    
    Set AddressColumn = ColumnRelativeObjectRange("A1:D10", 2)
    
    ColumnCounts ("A1:K10")
           
End Sub


