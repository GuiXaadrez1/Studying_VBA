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

' lembrando que o método Worbook("Nome_Pasta_Trabalho").Activate
' irá funcionar apenas se a pasta de trabalho estiver aberta.
' E sendo trabalhado no momento ou seja, para a pasta que você está mexendo atualmente
' Se a pasta de trabalho estiver fechada, você precisará fornecer o caminho completo.
' Com: Workbooks.Open ("C:\Usuários\JoãoPC2\Downloads\Pasta2.xlsm")

Private Function UsingWorkbookOpenActivete(ByVal nameWorkBook As String) As Workbook
    
    '   On Error ->  é a principal forma de tratamento de "Erro em Tempo de Execução" do VBA, informando o VBA o que fazer quando um erro desse tipo for encontrado.
    '   On Error pode ser utilizado de três maneiras:
    '
    '   On Error GoTo <Rótulo> -> Caso ocorra um erro, a execução irá pular para uma linha definida, geralmente ao final do código
    '   On Error Resume Next -> ara continuar a execução mesmo em caso de erro
    '   On Error GoTo 0 ->  é a opção padrão do VBA. Em caso de erro, a execução será interrompida e uma a mensagem de erro aparecerá
    
    On Error GoTo tratamentoError
        
        Dim WorkBookName As String
        Dim wb As Workbook          ' <-- ADICIONADO: você precisa de uma variável Workbook
        
        ' *** CORREÇÃO: Workbooks(nameWorkBook).Activate NÃO RETORNA OBJETO ***
        Set wb = Workbooks(nameWorkBook)  ' pega o Workbook da coleção
        
        wb.Activate                         ' ativa — este método não retorna nada
        
        WorkBookName = wb.Name
        
        Set UsingWorkbookOpenActivete = wb  ' retorna o Workbook correto
        
        UsingWorkbookOpenActivete.Worksheets(1).Activate ' Redireciona para a primeira Sheet (planilha) da pasta principal
        
        If Not UsingWorkbookOpenActivete Is Nothing Then
            
            MsgBox "Foi Acessada e Ativada para Trabalho a Pasta de trabalho: " & WorkBookName
            
        End If

    Exit Function    ' <-- NECESSÁRIO para não executar o handler sem erro

tratamentoError:
        MsgBox "ocorreu o seguinte Error: " & Err.Description

End Function



Sub ManipulationObjectWorkbooksInCollectionWorkbooks()
    UsingWorkbookOpenActivete ("LagoSul_VBA.xlsx")
End Sub

