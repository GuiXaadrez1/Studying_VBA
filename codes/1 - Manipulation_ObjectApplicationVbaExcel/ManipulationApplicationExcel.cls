VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit ' Todas as vari?veis devem ser explicitamente declaradas e corretamente

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Essa fun??o Explica como funciona as propriedades Application.Workbooks.Name
' Este c?digo do objeto application retorna no nome da pasta de trabalho ativa, e pode ser bastante ?til
' com intera??es com a fun??o msgbox.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ReturnWorkBookActiveWithApplicationCollection(ByVal nameWorkbook As String) As Workbook
    
    On Error Resume Next
        Set ReturnWorkBookActiveWithApplicationCollection = Application.Workbooks(nameWorkbook)
    On Error GoTo 0
    
    MsgBox ("Pasta de Trabalho Ativa ->" & " " & CStr(ReturnWorkBookActiveWithApplicationCollection.Name))

End Function


' Basicamente estamos fazendo a mesma coisa encima... Mas com Planilhas Ativas dentro do Workbook Ativo
' Application.ActiveSheet.Name

Private Function ReturWorkSheetActiveWithApplicationCollection(ByVal nameWorkbook As String) As Worksheet

    
    Dim wb As Workbook
    Dim nameSheet As String
    
    ' Reutiliza sua fun??o exatamente como est?
    Set wb = ReturnWorkBookActiveWithApplicationCollection(nameWorkbook)
    
    ' Agora sim: retorna a cole??o (Conjunto de Planilhas) de Sheets do Workbook
    ' Set ReturWorkSheetActiveWithApplicationCollection = wb.Worksheets
    
    ' Retorna a planilha ativa dentro do workbook
    Set ReturWorkSheetActiveWithApplicationCollection = wb.ActiveSheet
    
    If Not ReturWorkSheetActiveWithApplicationCollection Is Nothing Then
        MsgBox ("Planilha Ativa na " & CStr(wb.Name) & " -> " & ReturWorkSheetActiveWithApplicationCollection.Name)
    End If

End Function


Private Function ExitExcelWithApplicationCollection()
    
    ' Este ? um tipo de aplica??o utilizada quando queremos programar que a pasta de trabalho
    ' feche automaticamente ao t?rmino de determinada rotina.
    
    Application.Quit
    
End Function


Private Function DesactiveOperationTheWorkBookExcelSpecific()
    ' O c�digo application.interactive = true desabilita todas as opera��es de determinada pasta de trabalho,
    ' desta forma tenha sempre muito cuidado ao utilizar esta intera��o.
    Application.Interactive = True
End Function

Private Function ActivateOperationTheWorkBookExcelSpecific()
    Application.Interactive = False
End Function

Private Function HiddenModificationWithCode()
    ' Definir a propriedade ScreenUpdating como False para ocultar as altera��es enquanto elas s�o feitas via c�digo;
    Application.ScreenUpdating = False
End Function

Private Function ShowModificationWithCode()
    Application.ScreenUpdating = True
End Function

Private Function DisableFullScreen()
    ' O objeto application com o m�todo displayfullscreen permite retirar alguns elementos e deixar o Excel
    ' com a tela mais limpa, por�m n�o deixa a tela t�o limpa assim.
    Application.DisplayFullScreen = True
End Function

Private Function EnebleFullScreen()
    Application.DisplayFullScreen = True
End Function

Private Function DesableScrollBars()
    ' Este c�digo permite ocultar ou reexibir as duas barras de rolagens de uma vez,
    ' Seja ela: horizontal ou vertical
    ' e esta defini��o se d� pelo nome aplicado na propriedade true ou false.
    
    Application.DisplayScrollBars = False

End Function

Private Function EnableScrollBars()
    Application.DisplayScrollBars = True
End Function

Private Function DesableFormulaBar()
    ' Para interagir com a barra de f�rmulas ser� preciso utilizar o objeto application com o m�todo displayformulabar.
    Application.DisplayFormulaBar = False
End Function

Private Function EnableFormulaBar()
    Application.DisplayFormulaBar = True
End Function

Private Function DesableStatusBar()
    Application.DisplayStatusBar = False
End Function

Private Function EnableStatusBar()
    Application.DisplayStatusBar = True
End Function

Private Function SaveWorkBookActiveWithApplication(ByVal nameWorkbook As String)
    
    ' Vai funcionar perfeitamente desde que a Pasta de Trabalho Workbook esteja Ativa, ou seja, em execucao
    ReturnWorkBookActiveWithApplicationCollection(nameWorkbook).Save
    
End Function

Private Function DisableAlertsExcel()
    
    ' Application.DisplayAlerts controla se o Excel vai mostrar ou esconder as janelas de aviso enquanto o c�digo est� rodando.
    Application.DisplayAlerts = False

End Function


Private Function EnableAlertsExcel()

    Application.DisplayAlerts = True

End Function


Private Function DialogFindFile()
    
    ' O m�todo FindFile permite que o usu�rio selecione um arquivo a ser aberto no Excel;
    ' Esse arquivo vai ser outra pasta de trabalho
    Application.FindFile

End Function

Private Function DialogsWithApplicationCollection()
    
    ' O m�todo Dialogs exibir caixas de di�logos do Excel
    
    Application.Dialogs(XlBuiltInDialog.xlDialogOpen).Show

End Function

Private Function getFileDialog()

    ' O m�todo FileDialog obtem os nomes de arquivos e pastas;
    Application.FileDialog
    
End Function


Private Function inputBoxDialog()

    Dim nome As String
    
    nome = Application.InputBox( _
                Prompt:="Informe o seu nome:", _
                Title:="QUAL O SEU NOME", _
                Default:="seu nome")
    
    MsgBox "Seu nome �: " & nome
    
End Function



' Lembrando que a cole??o de objetos Application representa o pr?prio Excel em Si
' Ou seja, abrange todas as principais cole??es encadeadas
' Application := ( Workbooks, Sheets, Worksheets, Range, Columns, Cells )


Sub ManipulationAllObjectTheCollectionApplicationToExcel()
    
    
    DisableAlertsExcel
    
    ' EnableAlertsExcel
    
    HiddenModificationWithCode
     
    ReturnWorkBookActiveWithApplicationCollection ("LagoSul_VBA")
    
    ReturWorkSheetActiveWithApplicationCollection ("LagoSul_VBA")
    
    ' DesableFormulaBar
    
    ' EnableFormulaBar
    
    ' DesableScrollBars
     
    ' EnableScrollBars
     
    ' DesableStatusBar
        
    ' EnableStatusBar
    
    
    ' Salvando WorkBook Ativo
    ' Lembrando que o Workbook ativo � exatamente: A pasta ativa � sempre aquela selecionada no momento
    ' Desta forma equivale ao ActivateWorkbook.save
    ' Ele sempre salva apenas a que estiver em foco no momento, o workbook e o worksheet em foco
    
    ' Application.Workbooks.Application.ActiveWorkbook.Save
    
    ' DialogsWithApplicationCollection
    
    ' DialogFindFile
    
    ' getFileDialog
    
    ' InputBox
        
    inputBoxDialog
        
    SaveWorkBookActiveWithApplication ("LagoSul_VBA")
    
    MsgBox ("Planilha Salva com Sucesso!")
    
    ' ExitExcelWithApplicationCollection
    
    ShowModificationWithCode
    
End Sub



