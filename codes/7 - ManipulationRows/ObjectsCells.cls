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

' O objeto Rows é uma coleção horizontal, onde cada elemento é um Range que representa uma linha inteira.

' Formalmente:

' Rows(n) -> Range representando a linha n inteira.
' Range("B3:F20").Rows(2) -> a segunda linha interna desse intervalo.

' Conceitualmente:

' Onde Columns representa colunas verticais,
' Rows representa linhas horizontais contendo 1 × 16.384 células.

'Rows -> melhor para processamento sequencial, leitura horizontal, limpeza de registros.
'Columns -> melhor para estatística vertical e agrupamento de categorias.

'


Sub ManipulationObjectRowsWithVBA()
    ' Aqui não tem muito o que ver... vai seguir a mesma lógica que Range, Columns e Cells
End Sub


