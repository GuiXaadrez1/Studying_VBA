' Criar uma código VBA que formate a coluna G de NCM da planilha principal'
' E insira parte do codigo em cada coluna especifica -> capitulo, posisao, etc...
'
' Depois vá para planilha ReducaoNCM, faca a mesma formatacao da planilha
' Por fim faca o cruzamento de NCMs por niveis!


                                                    ''''''''''''''''''''''''
                                                    ' Configuracoes Gerais '
                                                    ''''''''''''''''''''''''

' nome da nossa Planilha principal
Const nameSheetPlanOne As String = "Itens das NF-es Recebidas - Aut"

' Inicia o range na celula 4 -> Planilha Principal
Const indexCellFirstColumn As Long = 4 ' linha 4

Const indexColumnPlanOne As Long = 7 ' Column G

' nome da nossa Planilha de Reducao de produto conforme seu codigo Ncm
Const nameSheetPlanTwo As String = "ReducaoNCM"

' Inicia o range na celula 4 -> Planilha ReducaoNCM
Const indexCellFirstColumnReducaoNCM As Long = 2 ' linha 2

Const indexCellColumnReducaoNCM As Long = 1 ' Coluna A

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



' ====================================================================================================
' CONFIGURAÇÃO PARTE CRUZAMENTO CNPJ
' ====================================================================================================

Const SourceSheetName As String = "Itens das NF-es Recebidas - Aut" ' Planilha que irevemos fazer o Cruzamento
Const SourceColumn As String = "G" ' Coluna G ao qual possui os Ncms a serem percorridos, comparados e cruzados
Const OutputColumn As String = "M" ' Coluna M ao qual vamos inserir as reduções conforme as comparações
Const StartRowSource As Long = 4 ' celula de inicio da iteracao

Const ReductionSheetName As String = "ReducaoNCM"
Const ReductionCodeColumn As String = "A"
Const ReductionTaxColumn As String = "G"
Const StartRowReduction As Long = 2

Const SheetCName As String = "PlanilhaC"
Const Ignore9DigitsInSheetC As Boolean = True


Private Function FormatarNcmPlanOne( _
    Optional ByVal nameSheet As String = nameSheetPlanOne, _
    Optional indexStartRange As Long = indexCellFirstColumn, _
    Optional indexExecutionColumnPlanOne As Long = indexColumnPlanOne _
) As Range
    
    ' Planilha Principal
    Dim Sheet As Worksheet
    
    ' Representa a celula especifica da nossa range
    Dim cel As Range
    
    ' Celula apos passar por um processo de limpeza e tratamento
    Dim cleaned As String
    
    ' Materializnado um Objeto Regex
    Dim regex As New RegExp
    
    ' Materializando um objeto Range
    Dim rng As Range
    
    ' quantidade de caracteres que o codigo ncm possui nesta range, intervalo de celulas
    Dim qtd As Long
    
    ' Definindo a ultima linha do nosso intervalo de celula
    Dim lastRow As Long
    
    ' Usando a nossa planilha
    Set Sheet = Worksheets(nameSheet)
    
    ' Ativando essa planilha
    Sheet.Activate
    
    ' puxando a nossa ultima Linha!
    lastRow = Sheet.Cells(Sheet.Rows.Count, indexExecutionColumnPlanOne).End(xlUp).Row
    
    ' Intervalo dinamico conforme a configuracao
    
    Set rng = Sheet.Range(Sheet.Cells(indexStartRange, indexExecutionColumnPlanOne), Sheet.Cells(lastRow, indexExecutionColumnPlanOne))
    
    ' Regex para remover tudo que nao e numero
    ' Usamos o bloco With para evitar fica escrevendo o nome do objeto para acessar o metodo ou atributo
    With regex
        .Global = True
        .Pattern = "\D"
    End With
    
    ' Percorrendo nosso Intervalo de celulas e acessando uma celuala em especifico
    For Each cel In rng
        
        ' Limpando espaços em branco se a celula for diferente de vazio
        If Trim(cel.Value) <> "" Then
            
            ' Mantem somente numeros
            cleaned = regex.Replace(cel.Value, "")
            qtd = Len(cleaned)
            
            ' -------------------------------
            '    COMPLETAR COM ZEROS ? ESQ.
            ' -------------------------------
            
            If qtd = 8 Then
                ' Ja esta completo e mantem como esta
                cleaned = cleaned
            
            ElseIf qtd = 7 Then
                cleaned = "0" & cleaned
            
            ElseIf qtd = 6 Then
                cleaned = "00" & cleaned
            
            ElseIf qtd = 5 Then
                cleaned = "000" & cleaned
            
            ElseIf qtd = 4 Then
                cleaned = "0000" & cleaned
            
            ElseIf qtd = 3 Then
                cleaned = "00000" & cleaned
            
            ElseIf qtd = 2 Then
                cleaned = "000000" & cleaned
            
            ElseIf qtd = 1 Then
                cleaned = "0000000" & cleaned
            
            ElseIf qtd = 0 Then
                ' Celula vazia depois da limpeza e ignora
                GoTo Proximo ' Simboliza o Continue  do Python para Loopings...
            
            ElseIf qtd > 8 Then
                
                ' Manter apenas os ULTIMOS 8 digitos
                cleaned = Right(cleaned, 8)
            
            End If
            
            ' Distribuindo a célula com Offset
            
            
            ' Formata tipo numero para string na exibicao do Excel
            cel.NumberFormat = "@"
            
            ' Pegando as duas primeiras celulas
            cel.Offset(0, 1).Value = "'" & CStr(Left(cleaned, 2))
            
            cel.Offset(0, 2).Value = "'" & CStr(Mid(cleaned, 3, 2))
             
            cel.Offset(0, 3).Value = "'" & CStr(Mid(cleaned, 5, 2))
            
            
            cel.Offset(0, 4).Value = "'" & CStr(Mid(cleaned, 7, 1))
            
            ' Pegando o ultimo valor do cod ncm
            cel.Offset(0, 5).Value = "'" & CStr(Right(cleaned, 1))
            
            
            ' -------------------------------
            '  FORMATAR COMO XX.XX.XX.X.X
            ' -------------------------------
            
            
            cel.Value = _
                Left(cleaned, 2) & "." & _
                Mid(cleaned, 3, 2) & "." & _
                Mid(cleaned, 5, 2) & "." & _
                Mid(cleaned, 7, 1) & "." & _
                Right(cleaned, 1)
            
        End If
        
Proximo:
    Next cel
    
    MsgBox "NCMs formatados com sucesso! Distribuicao feita."
    
End Function



Private Function FormatarPlanilhaReducaoNcm( _
    Optional ByVal nameSheet As String = nameSheetPlanTwo, _
    Optional indexStartRangeReducaoNCM As Long = indexCellFirstColumnReducaoNCM, _
    Optional indexExecutionColumnReducaoNCM As Long = indexCellColumnReducaoNCM _
) As Range

        Dim Sheet As Worksheet
        
        Dim cel As Range
        Dim valueCel As String
        Dim match As Object
        
        Dim rngNCM As Range
           
        ' Definindo variaveis dos MATCHES
        Dim regexIdentificarServico As String
        Dim regexCapNcm As String
        Dim regexPosNcm As String
        Dim regexSubPosValue As String
        Dim regexItemValue As String
        Dim regexSubItemValue As String
                  
        ' Materializando REGEX
        Dim regexNcmNove As New RegExp
        Dim regexNcmOito As New RegExp
        Dim regexNcmSete As New RegExp
        Dim regexNcmSeis As New RegExp
        Dim regexNcmCinco As New RegExp ' Representam os codigos Genericos
        Dim regexNcmQuatro As New RegExp
        Dim regexNcmDois As New RegExp
        Dim regexNcmUm As New RegExp
    
        ' Definindo a ultima linha do nosso intervalo de celula
        Dim lastRow As Long
    
    ' Ativando a nossa planilha ReducaoNCM
    
    Set Sheet = Worksheets(nameSheet)
        
    Sheet.Activate
    
    ' Configuracao dos REGEX
    With regexNcmNove
        .Global = False
        .Pattern = "(\d{1})(\d{2})(\d{2})(\d{2})(\d{1})(\d{1})"
    End With
    
    With regexNcmOito
        .Global = False
        .Pattern = "(\d{2})(\d{2})(\d{2})(\d{1})(\d{1})"
    End With
    
    With regexNcmSete
        .Global = False
        .Pattern = "(\d{2})(\d{2})(\d{2})(\d{1})"
    End With
    
    With regexNcmSeis
        .Global = False
        .Pattern = "(\d{2})(\d{2})(\d{2})"
    End With
    
    With regexNcmCinco ' Lembrando que estes vao idenficiar os gen?ricos
        .Global = False
        .Pattern = "(\d{2})(\d{2})(\d{1})"
    End With
    
    With regexNcmQuatro
        .Global = False
        .Pattern = "(\d{2})(\d{2})"
    End With
    
    With regexNcmDois
        .Global = False
        .Pattern = "(\d{2})"
    End With
    
    With regexNcmUm
        .Global = False
        .Pattern = "(\d{1})"
    End With
    
    ' Definindo a ultima linha
    lastRow = Sheet.Cells(Sheet.Rows.Count, indexExecutionColumnReducaoNCM).End(xlUp).Row
    
    ' Definindo o Range/Intervalo da coluna conforme a configuracao
    Set rngNCM = Sheet.Range(Sheet.Cells(indexStartRangeReducaoNCM, indexExecutionColumnReducaoNCM), Sheet.Cells(lastRow, indexExecutionColumnReducaoNCM))
    
    ' Percorrendo cada celula
    For Each cel In rngNCM
        
        'valueCel = Trim(cel.Value)
        valueCel = cel.Value
        valueCel = Trim(valueCel)
        valueCel = Replace(valueCel, ".", "")
        valueCel = Replace(valueCel, ",", "")
        valueCel = Replace(valueCel, " ", "")
        
        ' -------------------------------------------------------
        ' 9 caracteres ? formato completo
        ' -------------------------------------------------------
        
        If valueCel <> "" And Len(valueCel) = 9 Then
            
            If regexNcmNove.Test(valueCel) Then
                Set match = regexNcmNove.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
            
            regexIdentificarServico = match.SubMatches(0)
            regexCapNcm = match.SubMatches(1)
            regexPosNcm = match.SubMatches(2)
            regexSubPosValue = match.SubMatches(3)
            regexItemValue = match.SubMatches(4)
            regexSubItemValue = match.SubMatches(5)
            
            cel.Offset(0, 1).Value = "Servico nao faz parte do Produto ou seja Sem NCM"
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""
            
            valueCel = regexIdentificarServico & "." & regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue & "." & regexSubItemValue
            
            cel.Value = CStr(valueCel)
            
            
        ' -------------------------------------------------------
        ' 8 caracteres ? formato completo
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 8 Then
        
            If regexNcmOito.Test(valueCel) Then
                Set match = regexNcmOito.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            regexItemValue = match.SubMatches(3)
            regexSubItemValue = match.SubMatches(4)
            
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue & "." & regexSubItemValue

            cel.Offset(0, 1).Value = "'" & CStr(regexCapNcm)
            cel.Offset(0, 2).Value = "'" & CStr(regexPosNcm)
            cel.Offset(0, 3).Value = "'" & CStr(regexSubPosValue)
            cel.Offset(0, 4).Value = "'" & CStr(regexItemValue)
            cel.Offset(0, 5).Value = "'" & CStr(regexSubItemValue)
            
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
            
        ' -------------------------------------------------------
        ' 7 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 7 Then
        
            If regexNcmSete.Test(valueCel) Then
                Set match = regexNcmSete.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            regexItemValue = match.SubMatches(3)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = "'" & regexSubPosValue
            cel.Offset(0, 4).Value = "'" & regexItemValue
            cel.Offset(0, 5).Value = ""
            
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue & "." & regexItemValue
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
        ' -------------------------------------------------------
        ' 6 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 6 Then
        
            If regexNcmSeis.Test(valueCel) Then
                Set match = regexNcmSeis.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If

            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            regexSubPosValue = match.SubMatches(2)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = "'" & regexSubPosValue
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""
                
            valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue
            
            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
        
        ' -------------------------------------------------------
        ' 5 caracteres -> Representando os gen?ricos
        ' -------------------------------------------------------
        
        ElseIf valueCel <> "" And Len(valueCel) = 5 Then
        
            If regexNcmCinco.Test(valueCel) Then
                    Set match = regexNcmCinco.Execute(valueCel)(0)
                Else
                    GoTo TrataElse
                End If

                regexCapNcm = match.SubMatches(0)
                regexPosNcm = match.SubMatches(1)
                regexSubPosValue = match.SubMatches(2)
            
                cel.Offset(0, 1).Value = "'" & regexCapNcm
                cel.Offset(0, 2).Value = "'" & regexPosNcm
                cel.Offset(0, 3).Value = "'" & regexSubPosValue
                cel.Offset(0, 4).Value = ""
                cel.Offset(0, 5).Value = ""
                
                valueCel = regexCapNcm & "." & regexPosNcm & "." & regexSubPosValue
            
                ' Formata tipo n?mero para string na exibi??o do Excel
                cel.NumberFormat = "@"
            
                cel.Value = CStr(valueCel)
                
        ' -------------------------------------------------------
        ' 4 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 4 Then
            
            ' Eu Guilherme Henrique n?o sei que Bruxaria ? Essa!
            'valueCel = Replace(valueCel, ",", ".")
            'valueCel = Replace(valueCel, ".", "")
            
            If regexNcmQuatro.Test(valueCel) Then
                Set match = regexNcmQuatro.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            regexPosNcm = match.SubMatches(1)
            
            
            valueCel = regexCapNcm & "." & regexPosNcm
            
           ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
            
            cel.Value = CStr(valueCel)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = "'" & regexPosNcm
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

        ' -------------------------------------------------------
        ' 2 caracteres
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 2 Then
        
            If regexNcmDois.Test(valueCel) Then
                Set match = regexNcmDois.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

            cel.Value = CStr(valueCel)


        ' -------------------------------------------------------
        ' 1 caractere
        ' -------------------------------------------------------
        ElseIf valueCel <> "" And Len(valueCel) = 1 Then
            
            valueCel = "0" & valueCel
            
            If regexNcmDois.Test(valueCel) Then
                Set match = regexNcmDois.Execute(valueCel)(0)
            Else
                GoTo TrataElse
            End If
        
            regexCapNcm = match.SubMatches(0)
            
            cel.Offset(0, 1).Value = "'" & regexCapNcm
            cel.Offset(0, 2).Value = ""
            cel.Offset(0, 3).Value = ""
            cel.Offset(0, 4).Value = ""
            cel.Offset(0, 5).Value = ""

            ' Formata tipo n?mero para string na exibi??o do Excel
            cel.NumberFormat = "@"
        
            cel.Value = CStr(valueCel)
        
        Else
            GoTo TrataElse
        End If

        GoTo Proximo

TrataElse:
        cel.Offset(0, 1).Value = ""
        cel.Offset(0, 2).Value = ""
        cel.Offset(0, 3).Value = ""
        cel.Offset(0, 4).Value = ""
        cel.Offset(0, 5).Value = ""
    
        ' Formata tipo n?mero para string na exibi??o do Excel
        cel.NumberFormat = "@"
        
        cel.Value = CStr(valueCel)

Proximo:
    
    Next cel
    
    MsgBox "Planilha ReducaoNCM formatada com sucesso! Partes do codigo capturados com sucessos."


End Function




Sub FormatarPlanilhasCruzarDados()
    
    ' Materializando Variáveis para fazer a funcao de Cruzamento Funcionar
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsSrc As Worksheet: Set wsSrc = wb.Worksheets(SourceSheetName)
    Dim wsRed As Worksheet: Set wsRed = wb.Worksheets(ReductionSheetName)
    
     Application.ScreenUpdating = False
     
     Application.Interactive = False ' Impedindo do usuario mexer no excel para evitar erros inesperados
     
    Call FormatarNcmPlanOne
    
    Call FormatarPlanilhaReducaoNcm
    
    
    If CruzarNcmsPorNiveis(wsSrc, wsRed) Then
        '
    End If
    
    Application.ScreenUpdating = True
    Application.Interactive = True
    
    ' ATIVANDO A Planilha para Trabalho
    wsSrc.Activate
    MsgBox "Processamento concluído com sucesso!"

End Sub
