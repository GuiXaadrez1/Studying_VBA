VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub FormatarCPF_CNPJ_ColunaA_ServicosTomados()

    Dim regexCpf As New RegExp
    Dim regexCnpj As New RegExp
    Dim match As Object
    
    Dim linha As Long
    Dim valor As String
    Dim formatado As String

    regexCpf.Pattern = "(\d{3})(\d{3})(\d{3})(\d{2})"
    regexCpf.Global = False

    regexCnpj.Pattern = "(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})"
    regexCnpj.Global = False

    linha = 2

    Do While Cells(linha, 1).Value <> ""

        valor = Cells(linha, 1).Text
        
        valor = Replace(valor, ".", "")
        valor = Replace(valor, "-", "")
        valor = Replace(valor, "/", "")
        valor = Trim(valor)

        If Len(valor) <= 11 Then
            valor = String(11 - Len(valor), "0") & valor
        ElseIf Len(valor) <= 14 Then
            valor = String(14 - Len(valor), "0") & valor
        End If

        If Len(valor) = 11 Then
            
            Set match = regexCpf.Execute(valor)(0)
            
            formatado = match.SubMatches(0) & "." & _
                        match.SubMatches(1) & "." & _
                        match.SubMatches(2) & "-" & _
                        match.SubMatches(3)

            Cells(linha, 1).Value = formatado

        ElseIf Len(valor) = 14 Then

            Set match = regexCnpj.Execute(valor)(0)
            formatado = match.SubMatches(0) & "." & _
                        match.SubMatches(1) & "." & _
                        match.SubMatches(2) & "/" & _
                        match.SubMatches(3) & "-" & _
                        match.SubMatches(4)
            
            Cells(linha, 1).Value = formatado
        End If

        linha = linha + 1
    Loop

    MsgBox "CPF/CNPJ formatados com sucesso!"

End Sub

