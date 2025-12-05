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

' Basicamente essa fun��o vai adicionar uma nova planilha e ativar ela
Private Function AddNewSheetWithObjectCollectionWorkSheets() As Worksheet
    
    Set AddNewSheetWithObjectCollectionWorkSheets = Worksheets.add
    
    ' Ativando a planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheets.Activate
    
End Function

' Antes de uma planilha espec�fica, vamos criar ela e ativar
Private Function AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet(ByVal nameOtherSheet As String) As Worksheet
    
    ' Validando se a planilha n�o existe
    If Not SheetExists(nameOtherSheet) Then
    
        Err.Raise vbObjectError + 1000, , "A planilha '" & nameOtherSheet & "' n�o existe."
    
    End If
    
    ' Ap�s a valida��o, Criar a planilha CASO ELA N�O EXISTA
    
    Set AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet = Worksheets.add(Before:=Worksheets(nameOtherSheet))
    
    ' := -> � o operador de atribui��o nomeada de par�metros
    ' Basicamente, ele permite que voc� especifique qual argumento de um m�todo est� rescebendo um valor
    ' neste caso estamos passado o argumento nomeado "Before"
    ' Sem := o VBA n�o saberia a qual par�metro voc� est� se referindo, porque o m�todo .Add tem v�rios argumentos opcionais.
    ' Resumidamente, := -> Informar qual par�metro da fun��o est� sendo passado
    
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Name = "AddAntesPlanilha" & nameOtherSheet
    
    ' Ativando a Planilha para ser Trabalhada
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Activate
        
End Function

' Depois de uma planilha espec�fica, vamos criar ela e ativar

Private Function AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet(ByVal nameOtherSheet As String) As Worksheet

    If Not SheetExists(nameOtherSheet) Then
    
        Err.Raise vbObjectError + 1000, , "A planilha '" & nameOtherSheet & "' n�o existe."
    
    End If

    Set AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet = Worksheets.add(After:=Worksheets(nameOtherSheet))
     
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Name = "CriadoAposSheet" & nameOtherSheet
    
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Activate

End Function

' Depois da ultima planilha vamos criar uma outra e ativar a mesma
Private Function AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet() As Worksheet
    
    'Definindo uma vari�vel
    Dim ContSheets As Integer
    
    ContSheets = Worksheets.Count ' Essa propriedade vai retorar a quantidade total de planilhas nesta pasta de trabalho (WorkBook)
        
    ' Usando Set para referenciar o Objeto e ter acesso aos seus atributos, propriedades e metodos
    Set AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet = Worksheets.add(After:=Worksheets(ContSheets))
    
    ' Renomeando uma planilha
    AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet.Name = "AddAposUltimaSheet"
             
    ' Ativando a Planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet.Activate
    
End Function


' Criando uma Fun��o Auxiliar as Outras Fun��es

' ----------------------------------------------------------
' Verifica se uma planilha existe no Workbook atual.
' Aceita nome (String) ou �ndice num�rico.
' Retorna True se existir, False se n�o existir.
' ----------------------------------------------------------

Private Function SheetExists(ByVal sheetIdentifier As Variant) As Boolean
    
    ' Variant -> Tipo dedo que atende os dois casos.
    
    Dim ws As Worksheet ' Criando uma vari�vel que armazena um objeto Worksheet
    
    ' Definindo Tratamento de Error
    On Error Resume Next ' Se acontecer um erro na pr�xima linha, n�o pare o c�digo. Apenas continue
        
        Set ws = Worksheets(sheetIdentifier) ' Referenciando esse objeto Worksheet para adquirir e acessar todos seus atributos, metodos e propriedades
    
    On Error GoTo 0 ' Desliga o modo de ignorar erros. Volte ao comportamento normal do VBA.
    
    SheetExists = Not ws Is Nothing
    
End Function


Sub ExcutionFunctions()
    
    ' Agora vamos usar nossas fun��es!
    
    ' Adicionando Planilhs, rsrsrs, basicamente essa fun��o faz a mesma coisa que WorkSheets.add
    ' A unica diferen�a � que ele ativa a planilha para ser trabalhada
    AddNewSheetWithObjectCollectionWorkSheets ' Adicionar Planilha 2
    
    'AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet ("Planilha3")
    
    AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet ("TesteF1")
    
    AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet ("TesteF2")
    
    If SheetExists("TesteF1") And SheetExists("TesteF2") Then
        AddNewSheetWithObjectCollectionWorkSheetsAftherLastSheet
    End If
    
    ' Para liberar Espa�o na mem�ria -> Deletando Planilha
    'AddNewSheetWithObjectCollectionWorkSheetsBeforeOtherSheet.Delete
    'AddNewSheetWithObjectCollectionWorkSheetsAfterOtherSheet.Delete
         
End Sub
