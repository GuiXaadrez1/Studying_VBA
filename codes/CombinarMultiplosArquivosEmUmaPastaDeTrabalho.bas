VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Combina��o de Todas as Planilhas de Todas as Pastas de Trabalho Abertas em uma Nova Pasta de Trabalho como Planilhas Individuais

Option Explicit

Sub CombinarMultiplosArquivos()
On Error GoTo eh
'declarar vari�veis para conter os objetos necess�rios
   Dim wbDestination As Workbook
   Dim wbSource As Workbook
   Dim wsSource As Worksheet
   Dim wb As Workbook
   Dim sh As Worksheet
   Dim strSheetName As String
   Dim strDestName As String
'desativar a atualiza��o da tela para acelerar o processo
   Application.ScreenUpdating = False
'Primeiro, crie uma nova pasta de trabalho de destino
   Set wbDestination = Workbooks.Add
'obter o nome da nova pasta de trabalho para que voc� a exclua do loop abaixo
   strDestName = wbDestination.Name
'Agora, fa�a um loop em cada uma das pastas de trabalho abertas para obter os dados,
'mas exclua seu novo arquivo ou a pasta de trabalho de macro Pessoal
   For Each wb In Application.Workbooks
      If wb.Name <> strDestName And wb.Name <> "PERSONAL.XLSB" Then
         Set wbSource = wb
         For Each sh In wbSource.Worksheets
            sh.Copy After:=Workbooks(strDestName).Sheets(1)
         Next sh
      End If
   Next wb
'agora feche todos os arquivos abertos, exceto o novo arquivo e a pasta de trabalho da macro Personal.
   For Each wb In Application.Workbooks
      If wb.Name <> strDestName And wb.Name <> "PERSONAL.XLSB" Then
         wb.Close False
      End If
   Next wb

'remover a planilha um da pasta de trabalho de destino
   Application.DisplayAlerts = False
   Sheets("Planilha1").Delete
   Application.DisplayAlerts = True
'limpar os objetos para liberar a mem�ria
   Set wbDestination = Nothing
   Set wbSource = Nothing
   Set wsSource = Nothing
   Set wb = Nothing
'ativar a atualiza��o da tela quando conclu�da
   Application.ScreenUpdating = False
Exit Sub
eh:
   MsgBox Err.Description
End Sub
