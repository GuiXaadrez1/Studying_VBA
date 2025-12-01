Attribute VB_Name = "Módulo1"
Option Explicit

Sub matrizDinamicaConceitoBasico()
    Dim vldinamica() As String ' Perceba que não definimos uma quantidade fixa de tamanho, isso permite o excel preencher dinâmicamente e automaticamente conforme a necessidade
    
    ReDim vldinamica(4) 'ReDim -> (Redimensionar) é usada para alocar memória e definir o tamanho (dimensão) de uma matriz dinâmica.
    
    ' Observação: ReDim não preserva valores a cada execução, ele joga fora os armazenados!
    
    ReDim Preserve vldinamica(6) ' essa função Preserve faz o ReDim armazenar o novo numero de valores
End Sub


