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

Private Function AcessActivateSheetForIndexInCollectionWorSheet(ByVal indexSheet As Long) As Worksheet
    
    ' Define o objeto da função como o objeto da planilha
    Set AcessActivateSheetForIndexWithCollectionWorSheet = Worksheets(indexSheet) 'Perceba que estamos materialziando um WorkSheets pelo index
    
    AcessActivateSheetForIndexWithCollectionWorSheet.Activate
End Function


Private Function AcessActivateSheetForIndexWithCollectionSheet(ByVal indexSheet As Long) As Object

    AcessActivateSheetForIndexWithCollectionSheet = Sheets(indexSheet)
    
    AcessActivateSheetForIndexWithCollectionSheet.Activate
    
End Function


Sub diferrenceObjectWorksheetsAndSheets()

End Sub
