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


Private Function RerefenceStraightRange(ByVal intervalCell As String) As Object

    Set RerefenceStraightRange = Range(intervalCell)

End Function

Private Function RerefenceStraightRangeInWorksheetExplicit(ByVal nameSheet As String, intervalRange As String) As Range
    
    On Error Resume Next
        
        Set RerefenceStraightRangeInWorksheetExplicit = Worksheets(nameSheet)
        
        On Error GoTo 0 ' Desliga a manipulação de erro
    
            ' 2. Verifica se a atribuição falhou (ou seja, se a planilha não existe)
            If RerefenceStraightRangeInWorksheetExplicit Is Nothing Then
                
                ' Adiciona a nova planilha (o método Add retorna o novo objeto)
                Set RerefenceStraightRangeInWorksheetExplicit = Worksheets.Add
                
                ' Renomeando
                RerefenceStraightRangeInWorksheetExplicit.name = nameSheet
                
                ' Ativa a planilha assim que ela passa a existir
                RerefenceStraightRangeInWorksheetExplicit.Activate
                
                ' Aplica o renga rederenciado diretamente
                Set RerefenceStraightRangeInWorksheetExplicit = RerefenceStraightRangeInWorksheetExplicit.Range(intervalRange)
                
            Else
                ' Se a planilha existir
                RerefenceStraightRangeInWorksheetExplicit.Activate
                
                ' Atribui o Range ao retorno da função
                Set RerefenceStraightRangeInWorksheetExplicit = RerefenceStraightRangeInWorksheetExplicit.Range(intervalRange)
        End If
    
End Function


' Lembrando que esse Objeto Representa um Intervalo de Células
' Exemplo de intervalo de células -> "A1:A10", esse intervalo de célula é um Range aqui no VBA

Sub ManipulationObjectRangeInCollectionsRange()
     
    ' Basicamente definimos um intervalo de células diretamente nesta função bem como o valor da sua célula
    RerefenceStraightRange("A1:B10").Value = CStr("1000")
    
    
    RerefenceStraightRangeInWorksheetExplicit("Planilha2", "A1:B100").Value = CStr("10000")
    
End Sub

