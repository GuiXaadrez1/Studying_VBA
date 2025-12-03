Attribute VB_Name = "Módulo5"
Option Explicit

Sub ObjectRangeIntervalo()

' Esse objeto representa uma célula, uma coluna, uma linha, um conjunto de células, até todas as
' células de uma planilha. O objeto Range talvez seja o objeto mais utilizado no Excel VBA, e para
' ser usado deve-se referenciar a célula ou conjunto de células que será manipulada.


Range ("A1") ' Referência uma célula em específico
Range ("A2:B2") ' Referencia um intervalo de células
' Range ("Office") 'Referencia interavlo nomeado
Range ("c:c") ' referencia a coluna inteira
Range ("1:2") ' referencia linhas inteira
Range ("a1:b2,a10:c15") ' referencia intervalos não consecutivos, separa-se com vírgula

Worksheets("Planilha1").Range ("A5") ' referencia células de outras planilhas

End Sub
