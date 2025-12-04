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

' lembrando que o m�todo Worbook("Nome_Pasta_Trabalho").Activate
' ir� funcionar apenas se a pasta de trabalho estiver aberta.
' E sendo trabalhado no momento ou seja, para a pasta que voc� est� mexendo atualmente
' Se a pasta de trabalho estiver fechada, voc� precisar� fornecer o caminho completo.
' Com: Workbooks.Open ("C:\Usu�rios\Jo�oPC2\Downloads\Pasta2.xlsm")

Private Function UsingWorkbookOpenActivete(ByVal nameWorkBook As String) As Workbook
    
    '   On Error ->  � a principal forma de tratamento de "Erro em Tempo de Execu��o" do VBA, informando o VBA o que fazer quando um erro desse tipo for encontrado.
    '   On Error pode ser utilizado de tr�s maneiras:
    '
    '   On Error GoTo <R�tulo> -> Caso ocorra um erro, a execu��o ir� pular para uma linha definida, geralmente ao final do c�digo
    '   On Error Resume Next -> ara continuar a execu��o mesmo em caso de erro
    '   On Error GoTo 0 ->  � a op��o padr�o do VBA. Em caso de erro, a execu��o ser� interrompida e uma a mensagem de erro aparecer�
    
    On Error GoTo tratamentoError
        
        Dim WorkBookName As String
        Dim wb As Workbook          ' <-- ADICIONADO: voc� precisa de uma vari�vel Workbook
        
        ' *** CORRE��O: Workbooks(nameWorkBook).Activate N�O RETORNA OBJETO ***
        Set wb = Workbooks(nameWorkBook)  ' pega o Workbook da cole��o
        
        wb.Activate                         ' ativa � este m�todo n�o retorna nada
        
        WorkBookName = wb.name
        
        Set UsingWorkbookOpenActivete = wb  ' retorna o Workbook correto
        
        UsingWorkbookOpenActivete.Worksheets(1).Activate ' Redireciona para a primeira Sheet (planilha) da pasta principal
        
        If Not UsingWorkbookOpenActivete Is Nothing Then
            
            MsgBox "Foi Acessada e Ativada para Trabalho a Pasta de trabalho: " & WorkBookName
            
        End If

    Exit Function    ' <-- NECESS�RIO para n�o executar o handler sem erro

tratamentoError:
        MsgBox "ocorreu o seguinte Error: " & Err.Description

End Function


' Abrindo e ativando uma pasta de trabalho com Workboos.Open("nome_caminho_path_file")
Private Function UsingWorkBookOpenActivateOpeningWithWorkbooksCollection(ByVal pathFile As String) As Object
    
    On Error Resume Next
        
        ' Diferente do Anterior usamos para abrir a pasta Excel com um caminho espec�fico daquele arquivo...
        Set UsingWorkBookOpenActivateOpeningWithWorkbooksCollection = Workbooks.Open(pathFile)
         
End Function

' ThisWorkbook.Path � uma propriedade que retorna somente o caminho da pasta onde est� salvo o arquivo que cont�m o c�digo VBA.
Private Function ReturnPathFileWithVBA(ByVal nameArchiveExcel As String) As String
    
    ' Exemplo: Caminho Absoluto do arquivo -> C:\Users\Guilherme\Documents\Projetos\ControleEstoque.xlsm
    ' ThisWorkBook.Path vai retornar apenas -> C:\Users\Guilherme\Documents\Projetos
    ' Perceba que ele n�o coloca a barra "\" automaticamente no final do caminho.
    ' Para isso precisamos concatenar manualmente com &
    ' Exemplo: ThisWorkbook.Path & "\nomes.xlsx"
    
    ' OBSERVA��O -> � NECESS�RIO Salvar o arquivo que cont�m o c�digo
    ' PARA A PROPRIEDADE FUNCIONAR ADEQUADAMENTE
    
    On Error Resume Next
    
    MsgBox "PATH atual: " & ThisWorkbook.Path
    
    ReturnPathFileWithVBA = ThisWorkbook.Path & "\" & nameArchiveExcel

End Function


Private Function UsingPropertyThisWorkbookPathForOpeningWorkbooksWithVba(nameFileExcel As String) As Workbook
    
    On Error GoTo trata

        Set UsingPropertyThisWorkbookPathForOpeningWorkbooksWithVba = _
            Workbooks.Open(ReturnPathFileWithVBA(nameFileExcel))
    
        Exit Function

trata:
    MsgBox "Erro ao abrir o arquivo: " & Err.Description & vbCrLf & _
           "Caminho: " & ReturnPathFileWithVBA(nameFileExcel)

End Function


Private Function UsingVariablesTheEnvironmentForOpeningWorkbooks()
    
    ' Usamos a fun��o ou m�todo Environ("vari�vel de ambiente")
    ' justamente para acessar e usar uma vari�vel de ambiente espeficia do Host/Sistema Operacional (computador)

    UsingVariablesTheEnvironmentForOpeningWorkbooks = Environ("USERPROFILE") & "\Downloads\" & nameArchiveExcel
    
    ' Aqui construimos um caminho absoluto para acessar os arquivos excel que est�o na pasta Downloads

End Function

Sub ManipulationObjectWorkbooksInCollectionWorkbooks()
    
    ' UsingWorkbookOpenActivete ("LagoSul_VBA.xlsx")
    
    ' UsingWorkBookOpenActivateOpeningWithWorkbooksCollection ("C:\Users\guilherme.henrique\Downloads\LagoSul_VBA.xlsx")
  
    ' UsingPropertyThisWorkbookPathForOpeningWorkbooksWithVba ("LagoSul_VBA.xlsx")
    
End Sub

