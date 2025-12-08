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


Private Function AddItemInCollectionVBA(ByVal items As Variant) As Collection
    
    ' Sintaxe: obj.Add Item:=Nome, Key:=Codigo
    
    ' Criando uma vari�vel que j� materializa um Objeto Collection
    Dim MyFirstCollection As New Collection
     
    ' Fazendo uma atribui��o impl�cita de um Objeto Collection para a nossa funcao
    ' Usando a palavra chave de atribuicao de objetos Set (lembrando que funciona parecido com o This do chava)
    Set AddItemInCollectionVBA = MyFirstCollection
    
    ' Adicionando um Item a nossa Collection
    
    AddItemInCollectionVBA.Add Item:=items ' Usando o operador de atribuicao nomeado para mais clareza
    
    
End Function

Sub ManipulationCollectionVBA()

' Collection � uma estrutura de dados do pr�prio VBA, implementada como uma lista encadeada interna, capaz de armazenar qualquer tipo de
' elemento (Variant, objetos, arrays, strings, n�meros�).
' Objetos collections nos permite armazenar diversos itens em uma variavel de forma estruturada
' Ou seja, nos permite contruir uma cole��o de dados em mem�ria!
' Eles cont�m v�rios itens semelhantes que podem ser facilmente acessados e manipulados, mesmo que haja um grande n�mero de
' itens dentro da cole��o.
' Lembrando que diferente de dicionarios, collection sao nativos do vba e  funcionam de forma parecida com value:key

    'On Error Resume Next
        
        ' Armazenando uma vari�vel que armazena uma Collection
        Dim MinhaCollection As Collection
        
        ' Declrando o nosso incrementador
        Dim i As Long
        
        ' Nosso incrementador para o For Each
        Dim Element As Variant
    
        Set MinhaCollection = AddItemInCollectionVBA("ArmazenadoStringNaCollection")
        
        '.Count retorna o total de elementos na nossa collection
        Debug.Print "Quantidade de elementos: " & MinhaCollection.Count
        
        '.Item() -> Podemos acessar os itens de cada elemento usando index numerio ou o nome do elemento(chave)
        Debug.Print "Foi inserido o seguinte elemento a nossa collection: " & CStr(MinhaCollection.Item(1))
        
        ' Inserindo mais elementos, items na nossa Collection
        
        MinhaCollection.Add ("String2")
        MinhaCollection.Add (100)
        
        
        ' Iterando sobre a nossa Collection -> Come�ando do Index2, pois j� estamos acessando o index 1 anteriormente
        ' Observacao -> Nao existe index 0 em Collection e em Dicionarios
        For i = 2 To MinhaCollection.Count
        
            ' Debug.Print "Quantidade de elementos: " & MinhaCollection.Count
        
            '.Item() -> Podemos acessar os itens de cada elemento usando index numerio ou o nome do elemento(chave)
            Debug.Print "Foi inserido o seguinte elemento a nossa collection: " & CStr(MinhaCollection.Item(i))
        
        Next

        ' Chamando a Sub de remo��o, passando a MinhaCollection preenchida
        RemoveItemCollectionVBA targetCollection:=MinhaCollection, key:=1 ' -> Uma forma de passar a sub rotina auxiliar com (Expected operator -> :=) Argumentos Nomeados
        
        ' Outra forma de passar a sub rotina sem precisar explicitar os par�metros
        ' RemoveItemInCollectionVBA MinhaCollection, 1 ' Remove o item no �ndice 1 ("ArmazenadoStringNaCollection")
        
        Debug.Print "--- Ap�s a Remo��o ---"
        Debug.Print "Nova Quantidade de elementos: " & MinhaCollection.Count
        
        ' Outra Forma de iterar � usando For Each
        
        For Each Element In MinhaCollection
            MsgBox Element
        Next Element
                
End Sub


' Criando uma Sub rotina da rotina principal -> ManipulationCollectionVBA()
' Ela Recebe a Collection por refer�ncia (ByRef) e nao retorna valor (principal diferen�a para function)
Private Sub RemoveItemCollectionVBA(ByRef targetCollection As Collection, ByVal key As Variant)

    ' Com isso: N�o criamos uma nova cole��o. Usamos a cole��o passada por refer�ncia no par�metro (targetCollection).
    
    ' O On Error � �til se o item que voc� tentar remover n�o existir.
    On Error Resume Next
    
    ' .Remove() -> Usamos para remover um elemento de uma Collection
    targetCollection.Remove key

    ' Desliga o manipulador de erro, se necess�rio
    On Error GoTo 0
    
End Sub


