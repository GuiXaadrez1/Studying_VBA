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

' Range.Find e Range.FindNext
' Esses métodos formam o motor de busca mais rápido do Excel, ignorando loops linha-a-linha.

' Sintaxe: Range.Find(What, After, LookIn, LookAt, SearchOrder, SearchDirection, MatchCase)

' Parâmetros importantes:

' What — texto ou número procurado

' LookIn — onde pesquisar

' xlValues , xlFormulas

' LookAt — tipo de correspondência

' xlWhole (igual exato)

' xlPart (contém)

' SearchOrder — direção

' xlByRows ou xlByColumns

' MatchCase — diferencia maiúsculas

Private Function FindDataInCollectionDataExcel(ByVal findValue As Variant)
    
    Dim table As Range
    
    Dim DataCelFindInTable As Range
    
    ' Célula que vai representar a primeira ocorrência encontrada pelo método .Find()
    ' Este método sempre retorna a primeira ocorrência
    Dim firstMatch As Variant
    
    ' Vai procurar todas as informações das células ao redor diferente de vazio
    Set table = Range("C5").CurrentRegion
    
    ' Correção: garantir LookIn
    Set DataCelFindInTable = table.Find(What:=findValue, LookAt:=xlWhole, LookIn:=xlValues)
    
    If DataCelFindInTable Is Nothing Then
        MsgBox "Nenhum resultado encontrado."
        Exit Function
    End If
    
    ' Retorna a primeira ocorrência -> valor achado...
    firstMatch = DataCelFindInTable.Address
        
    ' Criando um Loop Do While para pecorrer e encontrar as proximas células que possui o valor
    ' com o mesmo endereço na célula DataCelFindInTable
    Do
        
        ' Funciona como se fosse o print do python, console.log do Typescript/JavaScript ou echo do PHP
        Debug.Print "Encontrado em: " & DataCelFindInTable.Address
        
        ' Já o método .FindNext percorre as seguintes para não parar no primeiro resultado...
        ' Correção: referência certa da variável (tabela -> table)
        Set DataCelFindInTable = table.FindNext(DataCelFindInTable)
        
        ' ternario no VBA
        If DataCelFindInTable Is Nothing Then Exit Do

    ' Correção: garantir que não entra loop infinito
    Loop While DataCelFindInTable.Address <> firstMatch
    
    Exit Function
    
End Function


Sub ManipulationObjectRangeInCollectionsRangeFindDatas()
    
    ' Por boa prática, para chamar funcoes privadas, usa-se a palavra chave de chamada Call
    ' Call FindDataInCollectionDataExcel("100")
    
End Sub

