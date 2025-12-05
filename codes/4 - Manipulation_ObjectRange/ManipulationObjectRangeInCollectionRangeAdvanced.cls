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

Private Function UsingMethodCurrentRegion(ByVal interval As String)

    ' A propriedade Range.CurrentRegion representa todo o bloco de células contíguas ao redor da célula inicial,
    ' delimitado por linhas e colunas vazias.
    
    ' Funciona como se fosse: Ctrl + Shift + * -> basicamente seleciona toda a nossa coleção de dados, ou seja, o intervalo inteiro de uma tabela do excel
    
    Dim rg As Range
    Dim tabela As Range
      
    ' Célula dentro da tabela
    Set rg = Range(interval)
    
    ' CurrentRegion pega a "ilha de dados -> coleção de dados"
    Set tabela = rg.CurrentRegion
    
    ' rg.CurrentRegion captura automaticamente toda a área contígua, eliminando a necessidade de
    ' calcular última linha/coluna manualmente.
    
    tabela.Value = "10"
    
    MsgBox "A tabela detectada possui " & _
           tabela.Rows.Count & " linhas e " & _
           tabela.Columns.Count & " colunas."

End Function

' OBERVACOES
' TODOS OS PARAMETROS OPCIONAIS NO VBA SAO DECLARADOS APOS OS PARAMETROS NAO OPCIONAIS
' Tipos Variant são os únicos totalmente opcionais sem valor padrão
' Para outros tipos, você deve definir um valor padrão: Function Exemple(Optional numero As Long = 0)

Private Function InsertColumnsWithRangeObject(ByVal intervalColumn As String)
    Range(intervalColumn).EntireColumn.Insert
End Function

Private Function DeleteColumnsWithRangeObject(ByVal intervalColumn As String)
    Range(intervalColumn).Delete
End Function

Private Function InsertRowsWithRangeObject(ByVal cellSelected As String)
    ' Insere uma linha encima desta célula selecionada
    Range(cellSelected).EntireRow.Insert
End Function

Private Function InsertMultipleRowsWithRangeObject(intervalRows As String)
    Range(intervalRows).Insert
End Function

Private Function DeleteRowsSpecific(intervalRows As String)
    Range(intervalRows).Delete
End Function

Private Function insertWithDeslocationUsingShift(ByVal intervalColumn As String) As Range
    
    On Error Resume Next
    
        ' empurra células da coluna B para baixo
        Set insertWithDeslocationUsingShift = Range(intervalColumn).Insert(Shift:=xlShiftDown)
        
        ' Shift:=xlShiftToRight ->  empurra células da linha "x" para direita

End Function

Sub ManipulationObjectRangeInCollectionsRangeAdvanced()
    
    ' Ela basicamente vai detectar uma Matriz por completo não pegando linhas e colunas
    ' UsingMethodCurrentRegion ("A1:H10")
     
    ' Insere uma coluna neste intervalo de célula e desloca as demais para direita
    ' InsertColumnsWithRangeObject ("A:A")
    
    ' InsertRowsWithRangeObject ("A3")
    
    'InsertMultipleRowsWithRangeObject ("5:10")
    
    ' DeleteRowsSpecific ("3:10") ' Estamos eliminando a linha A3 criada
    
    ' Essa função  basicamente vai inserir uma célula e empurrar as demais para baixo
    insertWithDeslocationUsingShift ("B2")
    
    ' Vai literalmente Deletar
    ' DeleteColumnsWithRangeObject ("A:A")
    
End Sub

