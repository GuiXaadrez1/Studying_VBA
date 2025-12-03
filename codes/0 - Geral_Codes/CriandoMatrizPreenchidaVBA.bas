VERSION 1.0 CLASS

BEGIN
  MultiUse = -1  'True
END

Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit ' Lembrando que serve para declarar todas as vari�veis adequadamente

Sub MatrizVBA()

    ' Primeiro Define o Tamanho do Vetor Multidimensional (Matriz)
    ' Segundo define a Dimensionalidade da matriz (linha X coluna = uma linha e duas colunas )
    Dim vetor2D(1 To 3, 1 To 2) As String
    
    ' linha 1, coluna 1
    vetor2D(1, 1) = "Office"
    
    ' linha 1, coluna 2
    vetor2D(1, 2) = "Power Point"
     
    ' linha 3, coluna 1
    vetor2D(2, 1) = "Excel"
    
    ' linha 2, coluna 2
    vetor2D(2, 2) = "One Driver"
    
    ' linha 2, coluna 2
    vetor2D(3, 1) = "One Driver"
    
    ' linha 3, coluna 2
    vetor2D(3, 2) = "One Driver"
    
    ' Preenchendo c�lulas com os valores definidos em nossa Matriz de Duas Dimens�es
    
    Range("A1") = vetor2D(1, 1)
    Range("B1") = vetor2D(1, 2)
    Range("A2") = vetor2D(2, 1)
    Range("B2") = vetor2D(2, 2)
    Range("A3") = vetor2D(3, 1)
    Range("B3") = vetor2D(3, 2)
    
    MsgBox "Foram preenchidas todos os valores da Matriz"
    
End Sub
