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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    CONFIGURAÇÃO PARA A FUNÇÃO CombinationRangeObjectAndCellsObject ' Com constantes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const rowOne As Integer = 1

Const columnOne As Integer = 1

Const rowTwo As Integer = 10

Const columnTwo As Integer = 1

Const valueCell As String = "100"


' OBSERVAÇÕES SOBRE FUNÇÕES:

' Função com 2+ argumentos -> VBA exige := ou Set

Private Function RerefenceStraightRange(ByVal intervalCell As String) As Object

    Set RerefenceStraightRange = Range(intervalCell)

End Function

Private Function RerefenceStraightRangeInWorksheetExplicit(ByVal nameSheet As String, intervalRange As String) As Range
    
    ' Declara uma variável auxiliar para a Planilha (Tipo WORKBOOK)
    Dim ws As Worksheet
    
    ' Tenta encontrar a planilha e usa On Error para lidar com a ausência
    On Error Resume Next
        Set ws = Worksheets(nameSheet)
    
    On Error GoTo 0 ' Se acontecer algum error, quebra a execução e aponta...

    ' Se a Planilha NÃO FOI ENCONTRADA (Is Nothing)
    If ws Is Nothing Then
        
        ' Cria e renomeia a nova planilha
        Set ws = Worksheets.Add
        ws.name = nameSheet
        
    End If
    
    ' A planilha (ws) AGORA EXISTE e está definida.
    ws.Activate
    
    ' O retorno da função (RerefenceStraightRangeInWorksheetExplicit)
    ' é explicitamente definido como o Range da planilha que existe/foi criada.
    Set RerefenceStraightRangeInWorksheetExplicit = ws.Range(intervalRange)
    
End Function


Private Function ReferenceWithCells(ByVal Row As Integer, Column As Integer) As Object

' Lembrando que Cells(linha, coluna) é um objeto que representa as células do Excel
' Os parâmetros são sempre  o index..., exemplo: 1 == linha 1, 1 na column == A

    Set ReferenceWithCells = Cells(Row, Column)
    
End Function


Private Function CombinationRangeObjectAndCellsObject(ByVal nameSheet As String) As Range
    
    Dim ws As Worksheet
    
    On Error GoTo Err
        
        ' Materializando um OBjeto WorkSheet
        Set ws = Worksheets(nameSheet)
        
        ' Ativando a planilha
        ws.Activate
        
        ' Se esse objet for diferente de Nada, então executar a nossa funcionalidade
        If Not ws Is Nothing Then
        
            Set CombinationRangeObjectAndCellsObject = ws.Range(ws.Cells(rowOne, columnOne), ws.Cells(rowTwo, columnTwo))
            
            CombinationRangeObjectAndCellsObject.Value = valueCell
            
            ' Sai com sucesso caso dê tudo certo! é uma boa prática coloca isso em funções.
            Exit Function
            
        End If
Err:
    ' Informa o erro
    MsgBox "Erro ao obter o Range: " & vbCrLf & _
           "Erro " & Err.Number & ": " & Err.Description, vbCritical
    
    ' Em caso de erro, a função retorna Nothing
    Set CombinationRangeObjectAndCellsObject = Nothing

End Function

Private Function DeslocationInterval(ByVal nameSheet As String, interval As String, valueCel As Variant) As Range
    
    Dim ws As Worksheet
    
    On Error GoTo Err
        
        Set ws = Worksheets(nameSheet)
        
        ws.Activate
        
        ' Offenset(0,1) ->  Desloca 0 linhas, 1 coluna (para a direita)
        Set DeslocationInterval = ws.Range(interval).Offset(0, 1)
        
        ' é possível redimensionar um Intervalo de células, usando o método Resize
        
        DeslocationInterval.Value = valueCel
        
        Exit Function

Err:
    
    MsgBox "Erro ao executar DeslocationInterval: " & vbCrLf & _
               "Erro " & Err.Number & ": " & Err.Description, vbCritical
               
    Set DeslocationInterval = Nothing
    
End Function


' Lembrando que esse Objeto Representa um Intervalo de Células
' Exemplo de intervalo de células -> "A1:A10", esse intervalo de célula é um Range aqui no VBA

Sub ManipulationObjectRangeInCollectionsRange()
      
    Dim incrementador As Integer
    Dim deslocarRange As Range
      
    ' Basicamente definimos um intervalo de células diretamente nesta função bem como o valor da sua célula
    ' RerefenceStraightRange("A1:B10").Value = CStr("1000")
    
    
    ' RerefenceStraightRangeInWorksheetExplicit("Planilha4", "A1:B100").Value = CStr("10000")
     
    ' Lembrando que esse := é uma atribuição aos parâmetros nomeados... Ou seja, escolhemos exatamente este parâmetro
    ' ReferenceWithCells(Row:=1, Column:=1).Value = 1000
    
    ' Iterando sobre células usando For
    
    'For incrementador = 1 To 10000
        'ReferenceWithCells(Row:=incrementador, Column:=1).Value = 1000
    'Next
    
    CombinationRangeObjectAndCellsObject ("Planilha1")
    
    ' Quando temos mais de um parâmetro na função do VBA é necessário usar
    ' Uma Referência de objeto para materialziar o Objeto retornado pela nossa função...
    Set deslocarRange = DeslocationInterval("Planilha1", "A1:A10", "100")
    
     
     
End Sub

