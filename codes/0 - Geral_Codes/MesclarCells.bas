' CRIAR UM VBA QUE EXECUTA O CALCULO DE CMS E IBS DE FORMA AUTOMATICA


'''''''''''''''''''''''''
' Configuaracoes Gerais '
'''''''''''''''''''''''''

Const sheetNamePattern As String = "Itens das NF-es Recebidas - Aut"


'Private Function CreatingTableCMSandIBS()

'End Function



Sub CalculateCBSandIBS()

    Dim sheetActive As Worksheet
    Dim rangeMerge As Range
    
    Set sheetActive = Worksheets(sheetNamePattern)
    
    Set rangeMerge = sheetActive.Range("AC1", "AC3")
    
    rangeMerge.Merge
    
    rangeMerge.Value = "Base Calculo"
    
    rangeMerge.HorizontalAlignment = xlCenter

End Sub
