Attribute VB_Name = "Módulo4"
Option Explicit
Sub ObjetoWorkSheet()

' O objeto Worksheet representa uma planilha, especificada da coleção Worksheets.

'  Worksheets. Esse objeto é também um membro da coleção Sheets.
' A coleção Sheets representa todas as planilhas da pasta de trabalho especificada.

' Worksheets("Planilha1").Visible = False ' habilita ou desabilita a visibilidade da planilha indicada da pasta de trabalho ativa.

Worksheets.Add ' Permite adicionar uma nova planilha na pasta de trabalho ativa

Worksheets("Planilha1").Activate  'Torna uma planilha indicada da coleção WorkSheets ativa.

Worksheets("Planilha6").Delete ' Exclui a planilha indicada da coleção WorkSheets

End Sub
