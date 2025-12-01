# Introdução 

A principal diferença é que a Function obrigatoriamente irá retornar um valor e deve ser chamada por algum outro procedimento

Funções ou Procedimentos como funções => Podem ser criadas novas funções utilizando o VBA, completamente adaptáveis
pelo usuário. Após a criação da função, esta pode ser utilizada igualmente às funções
predefinidas no excel. A sintaxe é a seguinte:

```vb
Option Explicit 

Function nome_funcao(Byval parametro_n1 As Integer, parametro n2 As Integer ) As Integer ' É possível declarar tipo de retorno na função e o tipo de retorno nos parâmetros
    ' ByVal É o mecanismo de passagem de valor para o parâmetro declarado na função

    n1 = 10
    n2 = 20

    nome_funcao = n1 + n2

End Function 



Sub somarDuasCelulas() ' lembrando que é possível passar parâmetros para um procedimento.

    Dim celAone As Range
    Dim celBobe as Range 

    celAone = Range("A1").Value
    celBone = Range("B2").Value

    resultado = CStr(nome_funcao(celAone, celBone))

    MsgBox "O resultado da soma dos valores foram:" & resultado

End Sub
```

### OBSERVAÇÕES

Não é possível executar uma Function sozinha. É necessário ser chamada por outro procedimento.

E.g. Sub sub_principal().

Uma vez definidas, as funções podem ser utilizadas normalmente pelo usuário no ambiente da planilha do Excel. Basta apenas chamar pelo nome da função, como com qualquer outra função pré-definida (SOMA, MÉDIA, etc.).