Attribute VB_Name = "Módulo2"
Option Explicit

' Agora vamos trabalhar um pouco com manipulação de Objetos, sendo os mais básicos:
' Application, Workbooks, Sheets, Range

' Vamos começar com Application


Sub applicationExcel()

    ' Objeto application -> é o excel em si, ele representa todo o aplicativo Microsoft Excel e seus derivados.
    ' Através dele, podemos configurar a visualização, execuções e outras funcionalidades do Excel.
    
    ' O Objeto Application possui várias propriedades e métodos:
    
    
    MsgBox Application.ActiveWorkbook.Name ' ActiveWorkbook é um método que retorna a pasta de trabalho ativa!
    
    MsgBox Application.ActiveSheet.Name ' ActiveSheet é um método que retorna a planilha ativa na pasta de trabalho
    
    ' Application.Quit -> Método que fecha as pastas de trabalho ativa
    
    
    Application.DisplayScrollBars = True ' Habilita e desabilita a barra de navegação entre as abas
    ' do excel. É necessário atribuir o valor verdadeiro (true) ou falso (false).
    
    Application.DisplayFormulaBar = True '  Habilita e desabilita a barra de fórmulas do excel e
    ' também precisa da atribuição de valor (true or false).

    
    Application.DisplayAlerts = False   ' Habilita e desabilita os avisos do excel e precisa da atribuição
    ' de valor(True Or False)
     
End Sub
