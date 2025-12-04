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
        
        WorkBookName = wb.name
        
        Set UsingWorkbookOpenActivete = wb  ' retorna o Workbook correto
        
        UsingWorkbookOpenActivete.Worksheets(1).Activate ' Redireciona para a primeira Sheet (planilha) da pasta principal
        
        If Not UsingWorkbookOpenActivete Is Nothing Then
            
            MsgBox "Foi Acessada e Ativada para Trabalho a Pasta de trabalho: " & WorkBookName
            
        End If

    Exit Function    ' <-- NECESSÁRIO para não executar o handler sem erro

tratamentoError:
        MsgBox "ocorreu o seguinte Error: " & Err.Description

End Function


' Abrindo e ativando uma pasta de trabalho com Workboos.Open("nome_caminho_path_file")
Private Function UsingWorkBookOpenActivateOpeningWithWorkbooksCollection(ByVal pathFile As String) As Object
    
    On Error Resume Next
        
        ' Diferente do Anterior usamos para abrir a pasta Excel com um caminho específico daquele arquivo...
        Set UsingWorkBookOpenActivateOpeningWithWorkbooksCollection = Workbooks.Open(pathFile)
         
End Function

' ThisWorkbook.Path é uma propriedade que retorna somente o caminho da pasta onde está salvo o arquivo que contém o código VBA.
Private Function ReturnPathFileWithVBA(ByVal nameArchiveExcel As String) As String
    
    ' Exemplo: Caminho Absoluto do arquivo -> C:\Users\Guilherme\Documents\Projetos\ControleEstoque.xlsm
    ' ThisWorkBook.Path vai retornar apenas -> C:\Users\Guilherme\Documents\Projetos
    ' Perceba que ele não coloca a barra "\" automaticamente no final do caminho.
    ' Para isso precisamos concatenar manualmente com &
    ' Exemplo: ThisWorkbook.Path & "\nomes.xlsx"
    
    ' OBSERVAÇÃO -> É NECESSÁRIO Salvar o arquivo que contém o código
    ' PARA A PROPRIEDADE FUNCIONAR ADEQUADAMENTE
    ' A recém-aberta pasta de trabalho sempre se tornará o ActiveWorkbook, permitindo-lhe interagir com ela facilmente.
    ' Assim podemos salvar específicamente aquela pasta de trabalho com Workbooks.Save
    
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
    
    ' Usamos a função ou método Environ("variável de ambiente")
    ' justamente para acessar e usar uma variável de ambiente espeficia do Host/Sistema Operacional (computador)

    UsingVariablesTheEnvironmentForOpeningWorkbooks = Environ("USERPROFILE") & "\Downloads\" & nameArchiveExcel
    
    ' Aqui construimos um caminho absoluto para acessar os arquivos excel que estão na pasta Downloads

End Function

Private Function SaveAllWorkbooks()

    Dim wb As Workbook   ' variável correta do loop

    ' Application.Workbooks pega todos os Workbooks abertos
    For Each wb In Application.Workbooks
        wb.Save
        MsgBox "Salvando a pasta: " & wb.name & vbCrLf & "Aperte OK para continuar..."
    Next wb

End Function


Private Function OpeningFilenameWithCollectionWoorkBook()
    
    Dim strArquivo As String
    
    ' Application.getOpenFilename -> é uma caixa de dialogo que é usado para o usuário abrir um arquivo
    ' Ele retorna o caminho absoluto daquele arquivo, pasta de trabalho que posteriormente vai ser aberto
    ' com o nosso Workbooks.Open()
    
    strArquivo = Application.GetOpenFilename()
    
    UsingWorkBookOpenActivateOpeningWithWorkbooksCollection (strArquivo)
    
End Function

Private Function AddNewWorkbooks()
    ' Essa propriedade cria uma nova pasta de trabalho
    ' Assim que ela é criada...  Se torna o ActiveWorkbook, lhe permitindo interagir com ela
    Workbooks.Add
End Function

Private Function CloseWorkBookAndSave()
    
    ' ActiveWorkbook.Close -> Fecha a pasta de trabalho Ativa
    '  Lembrando que vai salvar apenas um Workbook, que é justamente a que está ativa atualmente para trabalho
    ActiveWorkbook.Close SaveChanges:=True
    
    ' SaveChanges:=True -> Salva
    ' SaveChanges:=False -> Não salva
           
End Function


' Essa função vai simular a funcionalidade salvar como do excel! é possivel fazermos isso com workbooks
Private Function SaveWith()
    
    ' Para salvar a pasta de trabalho com um novo nome, no mesmo diretório, você pode simplesmente utilizar utilizar o código abaixo:
    ActiveWorkbook.SaveAs "NomeExemplo"
    
End Function

' Essa função vai simular a funcionalidade salvar como do excel! é possivel fazermos isso com workbooks
Private Function SaveWithDiretory(pathFileObsoluty As String)
    
    ' Para salvar a pasta de trabalho em um novo diretório, com uma extensão de arquivo específica, basta especificar o novo diretório e o nome do arquivo:
    
    ActiveWorkbook.SaveAs pathFileObsoluty
    
End Function


' uma fomra simples de como percorremos pastas de trabalhos ativas com VBA usando essa collection
Private Function LoopingInWoorkbooksActive()
    
    Dim wb As Workbook
    
    For Each wb In Workbooks
        MsgBox wb.name
    Next wb
    
End Function


Sub ManipulationObjectWorkbooksInCollectionWorkbooks()
    
    Application.ScreenUpdating = False
    
    ' UsingWorkbookOpenActivete ("LagoSul_VBA.xlsx")
    
    ' UsingWorkBookOpenActivateOpeningWithWorkbooksCollection ("C:\Users\guilherme.henrique\Downloads\LagoSul_VBA.xlsx")
  
    UsingPropertyThisWorkbookPathForOpeningWorkbooksWithVba ("LagoSul_VBA.xlsx")
    
    ' Referenciamento de pasta de Trabalho por Index... Funciona para Planilhas também rsrsrs
    ' Workbooks(1).Activate
    
    ' Função que abre uma caixa de diálogo para abrir uma pasta de trabalho
    OpeningFilenameWithCollectionWoorkBook
    
    AddNewWorkbooks
        
    'SaveAllWorkbooks
    
End Sub

