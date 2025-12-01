# Introdução
Basicamente ao criar um procedimento que é uma sequência de tarefas a aserem executadas e realizadas de forma automatizada, estamos criando uma MACRO com VBA

Declara o nome, argumentos e o código que formam o corpo de um procedimento Sub. Deve ser formatado seguindo estes parâmetros de sintaxe:

```vb
Sub nome_procedimento()

    ' Corpo do código, lógica de programação

End Sub
```

pós escrever seu código e nomeando o Sub é possível executá-lo, caso não tenha erros de compilação. 
Para isto existem as seguintes opções, clicando no botão play verde na barra de menus, usando o atalho F5, colocando o nome da Sub na janela de Verificação
Imediata ou mesmo utilizando a Depuração Passo a Passo F8. Caso o cursor esteja inserido
dentro do código, este será executado automaticamente. Caso contrário o software abrirá uma
tabela de macros para que o usuário selecione aquela a ser executada

```vb
Sub divide()

    | '<- imagina que isso seja o cusor do mouse  
   
    ActiveCell.Formula = _"=(" & Mid(ActiveCell.Formula,2)) & ")/100" 

End Sub
```

Neste caso, o cursor está situado dentro do código. Ao aplicar qualquer
um dos métodos mencionados acima, este será executado imediatamente.

Caso o cursor esteja fora do código, será aberta um interface que nos faz selecionar uma MACRO  a ser executada.
Escolha a macro que tem o nome do seu Sub ou procedimento...


## Caracteristicas

Uma Sub pode ser usada para executar outra Sub:

Exemplo:

```vb
    Option Explicit 

    Sub sub_principal() ' Sub que executa duas Subs
        sub_auxiliar1
        sub_auxiliar2
    End Sub
    
    Sub sub_auxiliar1()
        MsgBox "Uma Sub executando uma Sub"
    End Sub
    
    Sub sub_auxiliar2()
        MsgBox "Uma Sub executando mais outra Sub"
    End Sub
```

Sub também aceita um argumento ou multiplos argumentos:


```vb
' Sempre use o Option Explicit para codar em VBA para justamente todas as variáveis serem devidamente declaradas
Option Explicit 

' Sub com um argumento
Sub sub_principal()
    sub_argumento(10) '10 é um argumento
End Sub
Sub sub_argumento(x as Integer)
    MsgBox x
End Sub


' Sub com mais de um argumento -> Lembre-se de usar Call para poder usar uma sub com mais de dois argumentos
Sub sub_principal()
        Dim nota As Single
        Dim aluno As String

        nota=10
        aluno="Paulo"
        Call sub_argumento (nota, aluno) 'nota é um argumento, aluno é outro argumento
    End Sub
    Sub sub_argumento(n_prova as Single, nome as String)
        MsgBox nome & " tirou " & n_prova
    End Sub
```

Para passar mais de um argumento em uma Sub, é necessário utilizar a instrução Call antes do nome da Sub que se quer chamar.

Também é possível executar uma Sub com múltiplos argumentos omitindo a instrução de chamada Call e os parênteses:

sub_argumento nota, aluno


## Tipodes de Procedimentos (Sub)

Os tipos de Sub disponíveis são Private e Public:

    Public Sub: Cria uma subrotina pública que pode ser acessado por qualquer procedimento de qualquer módulo, e é exibida no Exibidor de Macros (Alt+F8).

    Private Sub: Cria uma subrotina privada que pode ser acessado apenas por procedimentos do mesmo módulo, e não é exibida no Exibidor de Macros (Alt+F8).



