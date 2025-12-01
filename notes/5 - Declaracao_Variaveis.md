## Utilização do Static no armazenamento de variáveis

No VBA as variáveis são descartadas ao final da execução de um código. Para que elas
sejam armazenadas existe a opção STATIC; essa opção armazena o valor final que a variável
assume e continua a execução do código, tornando a soma acumulativa.

```vb
Option Explicit 

Sub expiracao()
    Static x As Double

    x = x+1

    MsgBox x
End Sub
```

Na próxima execução o x terá o valor de dois, pois a variável foi armazenada e alocada na memória ram após o fim do código

## Declaração de constantes 

Podemos facilitar o trabalho no VBA por meio do uso de constantes. Elas são valores
fixos que, uma vez declarados, podem ser usados em todo o módulo (private) ou em toda
janela do VBE (public). Sua declaração pode ser feita dentro do Option Explicit, assim como a
declaração de variável, ou fora dele.


```vb
Option Explicit 

Sub Potencia_matematica()
    ' Criando a nossa constante, pode servir até como uma arquivo de configuração:

    Const potencial2 As integer = 2


    ' Variáveis que pode possuir valores mutáveis, ele é o nosso let do JavaScript

    Dim x As Double ' Armazena numeros de valores flutuantes
    Dim y As Long ' Armazena numeros inteiros muito grandes


    x = 10 

    resultado = x ^ potencial2

    MsgBox resultado
End Sub
```

## Função DATA e HORÁRIO

Essas duas funções têm declaração bem simples:

DATA = #mes/dia/ano#
HORÁRIO = #hora:min:segundo#

Para utilizá-las basta completar a data e o horário com os valores desejados. Lembrando que, tanto a data quanto o horário seguem padrões americanos, ou seja, a data 23

inverte mês e dia no CÓDIGO e o horário, para 17h por exemplo, o código automaticamente se altera para 5 PM.
