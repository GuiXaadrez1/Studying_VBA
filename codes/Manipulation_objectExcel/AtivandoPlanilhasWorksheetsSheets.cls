Option Explicit

' Usando WorkSheet para acessar e ativar uma planilha pelo nome
Private Function AcessActivateSheetForNameUsingCollectionWorkSheetObject(ByVal nameSheet As String) As Worksheet
    
    ' 1. Define o objeto da função como o objeto da planilha
    ' Lembrando que Set é como se fosse o This do Java... para referenciar Objetos a variáveis...
    Set AcessActivateSheetForNameUsingCollectionWorkSheetObject = Worksheets(nameSheet)
    
    ' 2. Ativa o objeto
    AcessActivateSheetForNameUsingCollectionWorkSheetObject.Activate
    
    ' Nota: O tipo de retorno deve ser Worksheet (o objeto)
End Function

' Usando Sheet para acessar e ativar uma planilha pelo nome
Private Function AcessActivateSheetForNameUsingCollectionSheetsObject(ByVal nameSheet As String) As Object
    
    ' 1. Define o objeto da função como o objeto da planilha (pode ser Worksheet ou Chart Sheet, então Object)
    Set AcessActivateSheetForNameUsingCollectionSheetsObject = Sheets(nameSheet)
    
    ' 2. Ativa o objeto
    AcessActivateSheetForNameUsingCollectionSheetsObject.Activate
    
    ' Nota: O tipo de retorno deve ser Object (pois Sheets pode conter diferentes tipos de objetos)

End Function

Sub diferrenceWorkSheetsForSheets()

' Diferença Entre as Collections :

' WorkSheet -> refere-se à coleção de todos os OBJETOS da planilha em uma pasta de trabalho, este não aceita gráficos como planilha

' Sheet -> refere-se a todos os tipos de PLANILHAS em uma pasta de trabalho, bem como planilhas de gráficos em uma pasta de trabalho, este aceita gráficos como planilha

' Se você tiver uma pasta de trabalho que tenha apenas planilhas e nenhuma planilha de gráfico, as coleções “Worsheets” e “Sheets” serão as mesmas.

' Por outro lado, quando temos uma ou mais planilhas de gráfico, a coleção “Sheets” será maior que a coleção “Worksheets”.

' Declarando variáveis

    Dim ws As Worksheet ' Variável que irá armazenar um Objeto Collection Worksheets
    Dim sht As Object ' Variável que irá armazenar um Object que pode ser tanto Char Sheet, WorkSheet, Sheet...
    
    Dim wsName As String
    Dim shtName As String
    
    ' *Renomeando a planilha recém-criada para que possa ser acessada
    With Worksheets.Add
        .Name = "Planilha2" ' <-- Define o nome da Planilha 1
    End With
    
    ' Planilha2**
    With Worksheets.Add
        .Name = "Planilha3" ' <-- Define o nome da Planilha 2
    End With
    
    MsgBox "Foram criadas e renomeadas as Planilhas 2 e 3!"
    
    wsName = "Planilha1"
    shtName = "Planilha2"
    
    ' Usar SET para atribuir o objeto retornado das Funções**
    Set ws = AcessActivateSheetForNameUsingCollectionWorkSheetObject(wsName)
    Set sht = AcessActivateSheetForNameUsingCollectionSheetsObject(shtName)
    
    ' Verificação de objeto (If Not Is Nothing)
    If Not ws Is Nothing Then
        MsgBox ("Foi Ativada a Planilha Worksheets: " & ws.Name)
    End If
    
    If Not sht Is Nothing Then
        MsgBox ("Foi Ativada a Planilha Sheets: " & sht.Name)
    End If
    
End Sub
