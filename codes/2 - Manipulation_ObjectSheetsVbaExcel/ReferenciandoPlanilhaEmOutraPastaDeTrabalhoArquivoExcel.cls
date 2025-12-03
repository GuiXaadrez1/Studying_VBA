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

' Private function -> : Indica que esta é uma função de uso interno, acessível apenas dentro do módulo onde foi declarada.

Private Function referenceOtherSheetWithObjectTheCollectionWorkSheets(ByVal workBookName As String, sheetName As String) As Worksheet
    
    ' Basicamente estamos referenciando uma outra planilha em outra pasta de trabalho ATIVA, iso é, em execução...
    ' Lembrando que Set, é o nosso This do Java em orientação a objeto
    ' Com ele podemos referenciar, acessar propriedades, atributos, metodos de Objetos materializados
    
    Set referenceOtherSheetWithObjectTheCollectionWorkSheets = Workbooks(workBookName).Worksheets(sheetName)
       
    referenceOtherSheetWithObjectTheCollectionWorkSheets.Activate
    
    ' Lembrando que a pasta de trabalho da planilha a ser ativada, acessada, deve estar aberta
    ' Mesmo que a pasta esteja em outra planilha, ao ativar a planilha desejada, o código nos leva exatamente para a aba da planilha

End Function

Sub ReferenceOtherSheet()
    
    ' Declaradno variável que aramazena, aloca um objeto Worksheet
    Dim OutraPlanilha As Worksheet
    
    Set OutraPlanilha = referenceOtherSheetWithObjectTheCollectionWorkSheets("Plano_Migraçao_Excel_v3.xlsx", "novo-myipac")
    
    If Not OutraPlanilha Is Nothing Then
        MsgBox ("Foi realizado o referenciamento de outra planilha e a mesma foi ativada em outra pasta de trabalho ativo" & OutraPlanilha.Name)
    End If
    
    ' Explicação Validação passo a passo
    
    ' Em VBA, quando você declara uma variável de objeto (como Dim OutraPlanilha As Worksheet), essa variável começa sem apontar para nada. Seu valor inicial é Nothing.
    
    ' Nothing: É um valor especial que indica que a variável de objeto não referencia nenhum objeto alocado na memória.
    
    ' Is: Este operador é usado apenas para comparar variáveis de objeto.
    ' Ele verifica se duas referências apontam para o mesmo objeto ou, neste caso, se um objeto é Nothing.
        
    ' Not: Nega o resultado lógico da comparação.
    
    ' Resumindo, Is Nothing faz uma comparação, se a variável nao esta apontando para Nothing, então retorna False
    ' Ai que entra o nosso not, pois vai negar o valor lógico retornado pela comparação de referencias de obejtos IS
    
End Sub



