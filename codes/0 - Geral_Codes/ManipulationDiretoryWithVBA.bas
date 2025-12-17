Option Explicit ' Obriga a declaração de variáveis, evitando erros de digitação

' Manipulando diretorios com as Class FileSystemObject do VBA
' FileSystemObject -> Lib nativa de manipulacao de diretorios do VBA


'''''''''''''''''''''''''''''
'   Configuracoes Gerais    '
'''''''''''''''''''''''''''''

Const SheetActiveName As String = "UNIFICADO"
Const patternPathDir As String = "Z:\CLIENTES ATIVOS"
Const patternPathSheetDir As String = "SOMA DAS NOTAS FISCAIS - MASTER FILIAL - 2025.xlsx"



Private Function GetFolderInDir(ByVal nameFolder As String) As String
    
    Dim fileSO As Object
    
    Set fileSO = CreateObject("Scripting.FileSystemObject")
    
    GetFolderInDir = CStr(fileSO.GetFolder(nameFolder))

End Function


Private Function ExistisFolder(ByVal pathDir As String) As Boolean

    Dim fileSO As Object
    
    Set fileSO = CreateObject("Scripting.FileSystemObject")
    
    If fileSO.FolderExists(pathDir) Then
        ExistisFolder = True
        
        Debug.Print ""
        Debug.Print "O diretório existe: " & pathDir
        
    Else
        ExistisFolder = False
        Debug.Print "O diretório não existe: " & pathDir
    End If
    
End Function

Private Function ListAllFilesTheMainDir(Optional ByVal pathDir As String = patternPathDir) As Variant
    
    Dim fileSO As Object
    Dim mainFolder As Object
    
    Dim subPath As Object
    Dim file As Object
     
    ' Variavel responsavel por obter o total de elementos em um diretorio
    Dim totalGeral As Long
    
    ' Criando matriz dinâmica para obter os resultados...
    Dim ResultsList() As Variant
    
    ' Criando contador para identificar qual pasta estamos acessando
    Dim Cont
    
    ' variaveis que vao obter o limite da matriz resultante
    Dim iRows As Long
    ' iColumns As Long
    
    
    ' Materializando os Objetos FileSystemObject
    
    Set fileSO = CreateObject("Scripting.FileSystemObject")
    
    Set mainFolder = fileSO.GetFolder(GetFolderInDir(pathDir))
    
    ' obtendo o total de elementos dos total de subdiretorios e o total de arquivos dentro da pasta base (main)
    totalGeral = mainFolder.Subfolders.Count + mainFolder.Files.Count
    
    If totalGeral > 0 Then
        
        ' Redimensionando a matriz para obter o tamanho extato para caber o total que existe no diretorio
        ReDim ResultsList(1 To totalGeral, 1 To 2)
    Else
        ListAllFilesTheMainDir = Empty
        Exit Function
    End If
    
    Cont = 0
    For Each subPath In mainFolder.Subfolders
        Cont = Cont + 1
        
        ' Preenchemos a matriz em memória
        ResultsList(Cont, 1) = Left$(subPath.Name, 3)     ' Nome da pasta (Ex: "649- AGFA"), no caso, apenas os tres primeiros digitos
        ResultsList(Cont, 2) = subPath.Path ' Caminho completo (Ex: "C:\Empresas\649- AGFA"), aqui obtemos o caminho completo
        
        ' Debug para acompanhar o progresso
        ' Debug.Print "Lendo pasta " & Cont & ": " & subPath.Name Left$(subPath.Name, 3)
        ' Debug.Print "Lendo pasta " & ": " & Left$(subPath.Name, 3) & " Caminho da Pasta: "; subPath.Path
    
    Next subPath
    
    ' Listando os Arquivos
    For Each file In mainFolder.Files
        Cont = Cont + 1
        
        ResultsList(Cont, 1) = file.Name ' Nome da pasta (Ex: "649- AGFA"), no caso, apenas os tres primeiros digitos
        ResultsList(Cont, 2) = file.Path ' Caminho completo
        
        ' Debug.Print "Arquivo dentro da pasta base: " & file.Name
    
    Next file
    
    
    ' Iterando sobre a matriz resultante para debugs
    'For iRows = LBound(ResultsList, 1) To UBound(ResultsList, 1)
        ' Exibe o Código (Coluna 1) e o Caminho/Nome (Coluna 2)
        'Debug.Print "Linha " & iRows & " -> Código: " & ResultsList(iRows, 1) & " | Info: " & ResultsList(iRows, 2)
    'Next iRows

    'Atribuindo a matriz preenchida ao retorno da função
    ListAllFilesTheMainDir = ResultsList
    
End Function

Private Function ListAllFilesTheSubDir(ByVal pathDir As String) As Variant
    
    Dim fileSO As Object
    Dim mainFolder As Object
    
    Dim subPath As Object
    Dim file As Object
     
    ' Variavel responsavel por obter o total de elementos em um diretorio
    Dim totalGeral As Long
    
    ' Criando matriz dinâmica para obter os resultados...
    Dim ResultsList() As Variant
    
    ' Criando contador para identificar qual pasta estamos acessando
    Dim Cont
    
    ' variaveis que vao obter o limite da matriz resultante
    Dim iRows As Long
    ' iColumns As Long
    
    
    ' Materializando os Objetos FileSystemObject
    
    Set fileSO = CreateObject("Scripting.FileSystemObject")
    
    Set mainFolder = fileSO.GetFolder(GetFolderInDir(pathDir))
    
    ' obtendo o total de elementos dos total de subdiretorios e o total de arquivos dentro da pasta base (main)
    totalGeral = mainFolder.Subfolders.Count + mainFolder.Files.Count
    
    If totalGeral > 0 Then
        
        ' Redimensionando a matriz para obter o tamanho extato para caber o total que existe no diretorio
        ReDim ResultsList(1 To totalGeral, 1 To 2)
    Else
        ListAllFilesTheMainDir = Empty
        Exit Function
    End If
    
    Cont = 0
    For Each subPath In mainFolder.Subfolders
        Cont = Cont + 1
        
        ' Preenchemos a matriz em memória
        ResultsList(Cont, 1) = subPath.Name     ' Nome da pasta (Ex: "649- AGFA"), no caso, apenas os tres primeiros digitos
        ResultsList(Cont, 2) = subPath.Path ' Caminho completo (Ex: "C:\Empresas\649- AGFA"), aqui obtemos o caminho completo
        
        ' Debug para acompanhar o progresso
        ' Debug.Print "Lendo pasta " & Cont & ": " & subPath.Name Left$(subPath.Name, 3)
        ' Debug.Print "Lendo pasta " & ": " & Left$(subPath.Name, 3) & " Caminho da Pasta: "; subPath.Path
    
    Next subPath
    
    ' Listando os Arquivos
    For Each file In mainFolder.Files
        Cont = Cont + 1
        
        ResultsList(Cont, 1) = file.Name ' Nome da pasta (Ex: "649- AGFA"), no caso, apenas os tres primeiros digitos
        ResultsList(Cont, 2) = file.Path ' Caminho completo
        
        ' Debug.Print "Arquivo dentro da pasta base: " & file.Name
    
    Next file
    
    ' Iterando sobre a matriz resultante para debugs
    'For iRows = LBound(ResultsList, 1) To UBound(ResultsList, 1)
        ' Exibe o Código (Coluna 1) e o Caminho/Nome (Coluna 2)
        'Debug.Print "Linha " & iRows & " -> FolderName/FileName: " & ResultsList(iRows, 1) & " | Info: " & ResultsList(iRows, 2)
    'Next iRows

    'Atribuindo a matriz preenchida ao retorno da função
    ListAllFilesTheSubDir = ResultsList
    
End Function


' Vamos colocar o valor do cod_interno como key e o valor vai ser um array, contendo Empresa e Regime Tributário
Private Function CollectionEmpFaturamento() As collection
    
    ' Criando Matriz Dinâmica
    Dim MatrizResults() As Variant
    Dim MatrizSubDirResults() As Variant
    
    ' Armazena a Ultima Linha do Array Resultante
    Dim lastRow As Long
    
    ' Dim iColumn As Long
    Dim iRows As Long
     
    Dim qtdCaracteres As Long ' Para Calcular a quantidade de caracteres do codigo...
     
    ' Atribuindo implicitamente a nova collection ao retorno da função
    Set CollectionEmpFaturamento = New collection
    
        
    lastRow = Cells(Rows.Count, "J").End(xlUp).Row

    
    ' Redimensiona do 7 até o fim para facilitar a vida
    ReDim MatrizResults(7 To lastRow, 1 To 3)
       
    ' transformando os tipos numerios em string...
    Range("C:C").NumberFormat = "@"
    
    ' Agora o UBound e LBound para obter os limites da Matriz Resultante para colocar na nossa collection
    For iRows = LBound(MatrizResults, 1) To UBound(MatrizResults, 1)
        
        ' Exemplo: imprimir o valor da célula na Janela de Verificação Imediata
        ' Debug.Print "Codigo Interno: "; Cells(iRows, "C") & " - " & "Empresa: "; Cells(iRows, "D") & " - "; "Regime Tributario: " & Cells(iRows, "J").Value
        
        qtdCaracteres = Len(CStr(Cells(iRows, "C").Value))
        
        If qtdCaracteres = 1 Then
            ' Se tem 1 dígito, coloca "00" na frente
            Cells(iRows, "C").Value = "00" & Cells(iRows, "C").Value
        ElseIf qtdCaracteres = 2 Then
            ' Se tem 2 dígitos, coloca "0" na frente
            Cells(iRows, "C").Value = "0" & Cells(iRows, "C").Value
        End If
        
        ' Armazena os valores na matriz primeiro
        MatrizResults(iRows, 1) = Cells(iRows, "C").Value ' Código Interno
        MatrizResults(iRows, 2) = Cells(iRows, "D").Value ' Empresa
        MatrizResults(iRows, 3) = Cells(iRows, "J").Value ' Regime
        
        CollectionEmpFaturamento.Add _
            Item:=Array(MatrizResults(iRows, 1), MatrizResults(iRows, 2), MatrizResults(iRows, 3)), _
            Key:=CStr(Cells(iRows, "C").Value)
    
    Next iRows
        
End Function

Public Sub ManipulationDir()
    
    ' vamos ativar a nossa planilha ao ser executada
    Dim ws As Worksheet
          
    'Dim collection As collection
    Dim vItem As Variant
    
    ' Matriz que vai obter o resultado da listagem de diretorio
    Dim ResultMatriz As Variant
    
    Dim MatrizSubDirResults As Variant
    
    ' Dim ResultSubDir As Variant
    
    ' variaveis que vao obter o limite da matriz resultante
    Dim iRows As Long
    'Dim iColumns As Long
    
    ' Criando Os segundos iteradores do segundo For de Procura
    Dim i As Long
    'Dim j As Long
    
    ' representa as oastas relativas
    Dim RelativeSubPath As String
    
    ' Criamos as duas variações possíveis para o ano atual
    Dim pathYearOptionalOne As String
    Dim pathYearOptionalTwo As String
    
    ' obtem o ano vigente!
    Dim currentYear As Integer
        
    Application.ScreenUpdating = False ' Desativado iteracoes do cod na planilha durante execucao do mesmo
    
     Application.DisplayAlerts = False ' Desativando as mensagens de alerta do excel
    
    ResultMatriz = ListAllFilesTheMainDir
    
    currentYear = Year(Date)
    
    For iRows = LBound(ResultMatriz, 1) To UBound(ResultMatriz, 1)
        
        ' Debug.Print ResultMatriz(iRows, 2)
        
        If ExistisFolder(ResultMatriz(iRows, 2)) = True Then
            
            ' apenas para quebrar linha
            Debug.Print ""
            Debug.Print ""
            
            MatrizSubDirResults = ListAllFilesTheSubDir(GetFolderInDir(ResultMatriz(iRows, 2)))
            
            RelativeSubPath = CStr(ResultMatriz(iRows, 2)) & "\DEPTO FISCAL"
            
            If ExistisFolder(RelativeSubPath) Then
                ' apenas para quebrar linha
                Debug.Print ""
                Debug.Print ""
            
                MatrizSubDirResults = ListAllFilesTheSubDir(RelativeSubPath)
                
                RelativeSubPath = CStr(ResultMatriz(iRows, 2)) & "\DEPTO FISCAL" & "\IMPOSTOS"
                
                If ExistisFolder(RelativeSubPath) Then
                    
                    pathYearOptionalOne = CStr(ResultMatriz(iRows, 2)) & "\DEPTO FISCAL" & "\IMPOSTOS" & "\" & currentYear
                    
                    pathYearOptionalTwo = CStr(ResultMatriz(iRows, 2)) & "\DEPTO FISCAL" & "\IMPOSTOS" & "\IMPOSTOS" & " " & currentYear
                    
                    ' Limpa a matriz para garantir que não use dados do cliente anterior
                    MatrizSubDirResults = Empty
                    
                    If ExistisFolder(pathYearOptionalOne) Then
                        
                        Debug.Print ""
                        
                         MatrizSubDirResults = ListAllFilesTheSubDir(pathYearOptionalOne)
                                            
                        For i = LBound(MatrizSubDirResults, 1) To UBound(MatrizSubDirResults, 1)
                            
                            If MatrizSubDirResults(i, 2) Like "*SOMA DAS NOTAS FISCAIS*" Then
                                
                                ' Workbooks.Open(MatrizSubDirResults(i, 2)).Activate
                                Workbooks.Open (MatrizSubDirResults(i, 2))
                                
                                Debug.Print ""
                                Debug.Print ""
                                
                                Debug.Print "Foi Acessado a pasta!"
                                
                                Exit For ' Encontrou? Para de procurar nesta pasta e volta para o loop de empresas
                            End If
                                                
                        Next i
                        
                    ElseIf ExistisFolder(pathYearOptionalTwo) Then
                        
                        Debug.Print ""
                        
                        MatrizSubDirResults = ListAllFilesTheSubDir(pathYearOptionalTwo)
                        
                        For i = LBound(MatrizSubDirResults, 1) To UBound(MatrizSubDirResults, 1)
                            
                            If MatrizSubDirResults(i, 2) Like "*SOMA DAS NOTAS FISCAIS*" Then
                                ' Workbooks.Open(MatrizSubDirResults(i, 2)).Activate
                                Workbooks.Open (MatrizSubDirResults(i, 2))
                                
                                Debug.Print ""
                                Debug.Print ""
                                
                                Debug.Print "Foi Acessado a pasta!"
                                
            
                                Exit For ' Encontrou? Para de procurar nesta pasta e volta para o loop de empresas
                            End If
                                                
                        Next i
                        
                    Else
                        Debug.Print "Pasta de " & currentYear & " não encontrada para este cliente."
                    End If
                    
                End If
                
            End If
                
            ' Exit For
            
        Else
            ' Exit For
        End If
        
        Debug.Print ResultMatriz(iRows, 2)
    
    Next iRows
    
    
    
    'Set ws = Worksheets(SheetActiveName)
    
    'ws.Activate
    
    ' CollectionEmpFaturamento()
     
    ' iterando sobre collection que possuem matriz como item
    'For Each vItem In CollectionEmpFaturamento()
        ' vItem(0) = Código, vItem(1) = Empresa, vItem(2) = Regime
        'Debug.Print "Cod (Key): " & vItem(0) & " | Empresa: " & vItem(1) & " | Regime: " & vItem(2)
    'Next vItem
    
End Sub
