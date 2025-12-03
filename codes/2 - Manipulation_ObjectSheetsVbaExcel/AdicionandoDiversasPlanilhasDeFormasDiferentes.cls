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

' Basicamente essa função vai adicionar uma nova planilha e ativar ela
Private Function AddNewSheetWithObjectCollectionWorkSheets() As Worksheet
    
    Set AddNewSheetWithObjectCollectionWorkSheets = Worksheets.add
    
    ' Ativando a planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheets.Activate
    
End Function

' Antes de uma planilha específica, vamos criar ela e ativar
Private Function AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet(ByVal nameOtherSheet As String) As Worksheet
    
    ' Validando se a planilha não existe
    If Not SheetExists(nameOtherSheet) Then
    
        Err.Raise vbObjectError + 1000, , "A planilha '" & nameOtherSheet & "' não existe."
    
    End If
    
    ' Após a validação, Criar a planilha CASO ELA NÃO EXISTA
    
    Set AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet = Worksheets.add(Before:=Worksheets(nameOtherSheet))
    
    ' := -> É o operador de atribuição nomeada de parâmetros
    ' Basicamente, ele permite que você especifique qual argumento de um método está rescebendo um valor
    ' neste caso estamos passado o argumento nomeado "Before"
    ' Sem := o VBA não saberia a qual parâmetro você está se referindo, porque o método .Add tem vários argumentos opcionais.
    ' Resumidamente, := -> Informar qual parâmetro da função está sendo passado
    
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Name = "AddAntesPlanilha" & nameOtherSheet
    
    ' Ativando a Planilha para ser Trabalhada
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Activate
        
End Function

' Depois de uma planilha específica, vamos criar ela e ativar

Private Function AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet(ByVal nameOtherSheet As String) As Worksheet

    If Not SheetExists(nameOtherSheet) Then
    
        Err.Raise vbObjectError + 1000, , "A planilha '" & nameOtherSheet & "' não existe."
    
    End If

    Set AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet = Worksheets.add(After:=Worksheets(nameOtherSheet))
     
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Name = "CriadoAposSheet" & nameOtherSheet
    
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Activate

End Function

' Depois da ultima planilha vamos criar uma outra e ativar a mesma
Private Function AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet() As Worksheet
    
    'Definindo uma variável
    Dim ContSheets As Integer
    
    ContSheets = Worksheets.Count ' Essa propriedade vai retorar a quantidade total de planilhas nesta pasta de trabalho (WorkBook)
        
    ' Usando Set para referenciar o Objeto e ter acesso aos seus atributos, propriedades e metodos
    Set AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet = Worksheets.add(After:=Worksheets(ContSheets))
    
    ' Renomeando uma planilha
    AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet.Name = "AddAposUltimaSheet"
             
    ' Ativando a Planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet.Activate
    
End Function


' Criando uma Função Auxiliar as Outras Funções

' ----------------------------------------------------------
' Verifica se uma planilha existe no Workbook atual.
' Aceita nome (String) ou índice numérico.
' Retorna True se existir, False se não existir.
' ----------------------------------------------------------

Private Function SheetExists(ByVal sheetIdentifier As Variant) As Boolean
    
    ' Variant -> Tipo dedo que atende os dois casos.
    
    Dim ws As Worksheet ' Criando uma variável que armazena um objeto Worksheet
    
    ' Definindo Tratamento de Error
    On Error Resume Next ' Se acontecer um erro na próxima linha, não pare o código. Apenas continue
        
        Set ws = Worksheets(sheetIdentifier) ' Referenciando esse objeto Worksheet para adquirir e acessar todos seus atributos, metodos e propriedades
    
    On Error GoTo 0 ' Desliga o modo de ignorar erros. Volte ao comportamento normal do VBA.
    
    SheetExists = Not ws Is Nothing
    
End Function


Sub ExcutionFunctions()
    
    ' Agora vamos usar nossas funções!
    
    ' Adicionando Planilhs, rsrsrs, basicamente essa função faz a mesma coisa que WorkSheets.add
    ' A unica diferença é que ele ativa a planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheets ' Adicionar Planilha 2
    
    'AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet ("Planilha3")
    
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet ("TesteF1")
    
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet ("TesteF2")
    
    If SheetExists("TesteF1") And SheetExists("TesteF2") Then
        AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet
    End If
    
    ' Para liberar Espaço na memória -> Deletando Planilha
    'AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Delete
    'AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Delete
         
End Sub
