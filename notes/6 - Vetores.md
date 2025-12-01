##  Vetores/Matriz unidimensional

Vetores, no código VBA, são armazenadores de valores de forma linear. É importante
lembrar que para vetores, o código VBA assume a posição zero sempre, ou seja, todo vetor
declarado tem a posição de número zero. Por exemplo, se você quer declarar um vetor de 3
posições você o fará das seguintes formas:


```vb
Option Explicit 

Sub vetor_String_Undimensional()

Dim vetor(1 to 3) As String ' Lembrando que criamos um vetor de tamnho 3, ou seja, vai de 0 a 2 no index

vetor(1) = "Excel"
vetor(2) = "Word"
vetor(3)= "LibreOffice"

End Sub
```

Os vetores podem servir para completar um intervalo dentro da planilha através do
próprio usuário por meio da função Range:


```vb
Option Explicit 

Sub vetor_String_Undimensional()

Dim vetor(3) ' de forma implícita fizemos a mesma coisa que está acima, porém vamos atribuir valores as células
Dim i As Long 

For i To 3 
    vetor(i) = InputBox("Nome do programa a ser armazenado?")
Next 

Range("A1") = vetor(1)
Range("A2") = vetor(2)
Range("A3") = vetor(3)

End Sub
```

## Matrizes

Matrizes são vetores bidimensionais ou tridimensionais, que existem no VBA.

As matrizes podem ser preenchidas pelo usuário através do InputBox também, da
mesma forma dos vetores. Podemos declarar uma matriz cúbica, com três dimensões
declarando 3 espaços onde, para duas dimensões, declarávamos dois

```vb
Sub vetor_Multidimensional()



End Sub
```

##  Matriz dinâmica

É uma matriz que tem seus espaços preenchidos de acordo com a necessidade. Ou
seja, o programador não delimita um intervalo de colunas e linhas.