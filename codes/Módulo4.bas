Attribute VB_Name = "M�dulo4"
Option Explicit
Sub ObjetoWorkSheet()

' O objeto Worksheet representa uma planilha, especificada da cole��o Worksheets.

'  Worksheets. Esse objeto � tamb�m um membro da cole��o Sheets.
' A cole��o Sheets representa todas as planilhas da pasta de trabalho especificada.

' Worksheets("Planilha1").Visible = False ' habilita ou desabilita a visibilidade da planilha indicada da pasta de trabalho ativa.

Worksheets.Add ' Permite adicionar uma nova planilha na pasta de trabalho ativa

Worksheets("Planilha1").Activate  'Torna uma planilha indicada da cole��o WorkSheets ativa.

Worksheets("Planilha6").Delete ' Exclui a planilha indicada da cole��o WorkSheets

End Sub
